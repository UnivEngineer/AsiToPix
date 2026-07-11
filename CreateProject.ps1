# Run as Admin for Symlinks
[CmdletBinding(SupportsShouldProcess = $true)]
param()

Set-StrictMode -Version Latest
$ErrorActionPreference = 'Stop'

$environmentModule = Join-Path $PSScriptRoot "src\AsiToPix.Environment.psm1"
Import-Module $environmentModule -Force
Initialize-AsiToPixEnvironment

$pathsModule = Join-Path $PSScriptRoot "src\AsiToPix.Paths.psm1"
Import-Module $pathsModule -Force

function Read-CreateProjectConfirmation {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Prompt,

        [bool]$DefaultYes = $true
    )

    $suffix = if ($DefaultYes) { "(Y/n)" } else { "(y/N)" }
    do {
        $answer = (Read-Host "$Prompt $suffix").Trim()
        if ([string]::IsNullOrWhiteSpace($answer)) {
            return $DefaultYes
        }

        $firstChar = $answer[0]
        if ($firstChar -in @([char]'y', [char]'Y', [char]0x0434, [char]0x0414)) { return $true }
        if ($firstChar -in @([char]'n', [char]'N', [char]0x043d, [char]0x041d)) { return $false }

        Write-Host "[!] Enter Y/N or the Cyrillic yes/no initials." -ForegroundColor Red
    } while ($true)
}

function Resolve-CreateProjectFullPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (Test-Path -LiteralPath $Path) {
        return (Resolve-Path -LiteralPath $Path).ProviderPath
    }

    return [System.IO.Path]::GetFullPath($ExecutionContext.SessionState.Path.GetUnresolvedProviderPathFromPSPath($Path))
}

function Test-CreateProjectPathInside {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ChildPath,

        [Parameter(Mandatory = $true)]
        [string]$ParentPath
    )

    $childFull = Resolve-CreateProjectFullPath -Path $ChildPath
    $parentFull = Resolve-CreateProjectFullPath -Path $ParentPath
    $separator = [System.IO.Path]::DirectorySeparatorChar
    $parentWithSeparator = if ($parentFull.EndsWith($separator)) {
        $parentFull
    } else {
        $parentFull + $separator
    }

    return $childFull.Equals($parentFull, [System.StringComparison]::OrdinalIgnoreCase) -or
        $childFull.StartsWith($parentWithSeparator, [System.StringComparison]::OrdinalIgnoreCase)
}

function New-CreateProjectDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (Test-Path -LiteralPath $Path -PathType Container) {
        return
    }

    if (Test-Path -LiteralPath $Path) {
        throw "Cannot create directory because a non-directory item already exists: $Path"
    }

    if ($script:PSCmdlet.ShouldProcess($Path, "Create directory")) {
        New-Item -ItemType Directory -Path $Path -Force -ErrorAction Stop | Out-Null
    }
}

function Remove-CreateProjectSourceDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,

        [Parameter(Mandatory = $true)]
        [string]$SetupRoot
    )

    if (-not (Test-Path -LiteralPath $SourcePath)) {
        return
    }

    $sourceItem = Get-Item -LiteralPath $SourcePath -Force
    if (-not $sourceItem.PSIsContainer -or $sourceItem.LinkType -eq 'SymbolicLink') {
        throw "Refusing to clean Source because it is not an ordinary directory: $SourcePath"
    }

    if ($sourceItem.Name -ne "Source") {
        throw "Refusing to clean a directory that is not named Source: $SourcePath"
    }

    if (-not (Test-CreateProjectPathInside -ChildPath $SourcePath -ParentPath $SetupRoot)) {
        throw "Refusing to clean Source outside the expected project root. Source: $SourcePath; Project root: $SetupRoot"
    }

    if ($script:PSCmdlet.ShouldProcess($SourcePath, "Remove existing project Source directory")) {
        Remove-Item -LiteralPath $SourcePath -Recurse -Force -ErrorAction Stop
    }
}

function New-CreateProjectSymbolicLink {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [string]$Target
    )

    if (Test-Path -LiteralPath $Path) {
        $existingItem = Get-Item -LiteralPath $Path -Force
        if ($existingItem.LinkType -ne 'SymbolicLink') {
            throw "Cannot create symbolic link because an ordinary item already exists: $Path"
        }

        $existingTarget = @($existingItem.Target)[0]
        if ($existingTarget -eq $Target) {
            return
        }

        throw "Symbolic link already exists with a different target. Path: $Path; Existing target: $existingTarget; New target: $Target"
    }

    $parentPath = Split-Path -Path $Path -Parent
    New-CreateProjectDirectory -Path $parentPath
    if ($script:PSCmdlet.ShouldProcess($Path, "Create symbolic link to '$Target'")) {
        New-Item -ItemType SymbolicLink -Path $Path -Value $Target -ErrorAction Stop | Out-Null
    }
}

function Copy-CreateProjectDirectory {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Source,

        [Parameter(Mandatory = $true)]
        [string]$Destination
    )

    if (Test-Path -LiteralPath $Destination) {
        $destinationItem = Get-Item -LiteralPath $Destination -Force
        if (-not $destinationItem.PSIsContainer -or $destinationItem.LinkType -eq 'SymbolicLink') {
            throw "Cannot copy files because the destination already exists and is not an ordinary directory: $Destination"
        }

        return
    }

    New-CreateProjectDirectory -Path $Destination
    if ($script:PSCmdlet.ShouldProcess($Destination, "Copy files from '$Source'")) {
        & robocopy $Source $Destination /E /MT /R:2 /W:5 | Out-Null
        if ($LASTEXITCODE -gt 7) {
            throw "Robocopy failed with exit code $LASTEXITCODE while copying '$Source' to '$Destination'."
        }
    }
}

Write-Host "--- ASIAir SMART SCANNER v29 ---" -ForegroundColor Cyan

$baseZ = Resolve-AstroPhotoRoot

# 1. AUTO-DETECT
$inputPath = (Read-Host "Paste path to lights folder (or any .fit file inside)").Trim('"')

# If a file path was given, use its parent folder
if ($inputPath -match '\.\w+$') { $inputPath = Split-Path $inputPath -Parent }

# Parse the path for project metadata
if ($inputPath -match 'ASIAir\\(?<obj>[^\\]+)\\(?<season>[^\\]+)\\(?<tel>.*?) @ (?<cam>ASI\d+[^\\]*)\\Good\\[^\\]+') {
    $parsedObj      = $Matches['obj']
    $parsedSeason   = $Matches['season']
    $parsedTel      = $Matches['tel']
    $parsedCamShort = $Matches['cam']
} else {
    Write-Host "[ERROR] Path parse error!" -ForegroundColor Red
    Write-Host "  Your path: $inputPath" -ForegroundColor Yellow
    Write-Host "  Expected:  ...\ASIAir\<Object>\<Season>\<Scope> @ <Cam>\Good\<Filter>\<Date>" -ForegroundColor DarkGray
    Write-Host "  Example:   <AstroPhotoRoot>\ASIAir\M 3\Lake\LX200 @ 0.63x @ ASI2600\Good\L\26.04.09" -ForegroundColor DarkGray
    Write-Host "  (Is the filter folder missing between 'Good' and the date folder?)" -ForegroundColor Red
    exit 1
}

# ── Interactive parameter confirmation ───────────────────────────────────────
function Confirm-Param {
    param([string]$Label, [string]$Detected)
    # Print label, arrow and detected value on one line, then prompt on same line
    $labelPad = $Label.PadRight(7)
    $detPad = $Detected.PadRight(25)
    Write-Host "  $labelPad -> " -ForegroundColor DarkGray -NoNewline
    Write-Host $detPad -ForegroundColor Yellow -NoNewline
    Write-Host "  >:" -ForegroundColor DarkGray -NoNewline
    $ans = Read-Host
    if ($ans -eq $null -or $ans.Trim() -eq "") { return $Detected } else { return $ans.Trim() }
}

Write-Host "`n-- Detected parameters -- (Enter to accept, or type override)" -ForegroundColor Cyan
# Print each parameter label and detected value; Confirm-Param will prompt on same line
Write-Host "" -NoNewline

$astroObj = Confirm-Param "Object" $parsedObj
$season   = Confirm-Param "Season" $parsedSeason
$telSetup = Confirm-Param "Scope"  $parsedTel

# camShort = path segment from folder name (e.g. ASI2600, may lack MM/MC suffix)
$camShort = $parsedCamShort

# Structure
$safeObj = $astroObj.Replace(" ", "_");

# Project Path confirmation (default: next to script)
$defaultProjectPath = Join-Path $PSScriptRoot $safeObj
$projectPath = Confirm-Param "Project" $defaultProjectPath
Write-Host ""

$safeSetup = "$($season)_$($telSetup)_$($camShort)".Replace(" ", "_").Replace("@","").Replace("__","_")
$setupRoot = Join-Path $projectPath $safeSetup
$sourcePath = Join-Path -Path $setupRoot -ChildPath "Source"
$pixPath = Join-Path -Path $setupRoot -ChildPath "Pix"

if (Test-Path -LiteralPath $sourcePath) {
    if (Read-CreateProjectConfirmation -Prompt "Clean [Source] folder?") {
        Remove-CreateProjectSourceDirectory -SourcePath $sourcePath -SetupRoot $setupRoot
    }
}
foreach ($p in @($sourcePath, $pixPath)) {
    New-CreateProjectDirectory -Path $p
}

Write-Host ""
$mergeFilters = Read-CreateProjectConfirmation -Prompt "Allow WBPP to merge sessions with filters None, IRC and Trib into a single one L?"

$asiairRoot = Join-Path -Path $baseZ -ChildPath "ASIAir"
$objectRoot = Join-Path -Path $asiairRoot -ChildPath $astroObj
$seasonRoot = Join-Path -Path $objectRoot -ChildPath $season
$setupSourceRoot = Join-Path -Path $seasonRoot -ChildPath "$telSetup @ $camShort"
$lightsRoot = Join-Path -Path $setupSourceRoot -ChildPath "Good"
if (-not (Test-Path -LiteralPath $lightsRoot -PathType Container)) {
    throw "Lights root not found: $lightsRoot"
}
$pendingLinks = @()
 # Camera mapping summary: collect observed raw filters, sanitized filters and targets per camera
 $camMap = @{}

# Filter mapping: raw filter name from filename -> normalized output filter + WBPP target group
# OSC (*MC): IRC/L/Trib -> L/RGB | None -> None/RGB | HO/UHC -> HO/HO | SO -> SO/SO
# Mono (*MM): L -> L/RGB | None -> None/RGB | R/G/B -> R,G,B/RGB | H/Ha -> H/H | S/SII -> S/S | O/OIII -> O/O
function Get-FilterMap($rawFilt, $camType) {
    $r = $rawFilt.Trim()
    if ($camType -eq "OSC") {
        switch -Regex ($r) {
            '^(IRC|Trib)$'         { return @{ Filter="L";    Target="RGB" } }
            '^L$'                  { return @{ Filter="L";    Target="RGB" } }
            '^None$'               { return @{ Filter="None"; Target="RGB" } }
            '^(HO|UHC)$'           { return @{ Filter="HO";   Target="HO"  } }
            '^SO$'                 { return @{ Filter="SO";   Target="SO"  } }
            default                { return @{ Filter=$r;     Target=$r    } }
        }
    } else {
        # Mono
        switch -Regex ($r) {
            '^L$'                  { return @{ Filter="L";    Target="L" } }
            '^None$'               { return @{ Filter="None"; Target="None" } }
            '^R$'                  { return @{ Filter="R";    Target="R" } }
            '^G$'                  { return @{ Filter="G";    Target="G" } }
            '^B$'                  { return @{ Filter="B";    Target="B" } }
            '^(H|Ha)$'             { return @{ Filter="H";    Target="H"   } }
            '^(S|SII)$'            { return @{ Filter="S";    Target="S"   } }
            '^(O|OIII|OII)$'       { return @{ Filter="O";    Target="O"   } }
            default                { return @{ Filter=$r;     Target=$r    } }
        }
    }
}

# Calibration Path Helper - $cb = calibration base for this session's camera
function Get-CalibPath($type, $gain, $rawTemp, $expValue, $cb) {
    $sub = switch($type) { "Darks" {"darks"} "Biases" {"biases"} "FlatDarks" {"flat-darks"} }
    $numTemp = [double]($rawTemp -replace '[^-\d\.]', '')
    $targetTempFolder = "$([Math]::Round($numTemp / 5.0) * 5)C"
    
    foreach ($m in @("Master", "Source")) {
        $base = Join-Path -Path $cb -ChildPath "$m\$sub\Gain$gain\$targetTempFolder"
        if (Test-Path -LiteralPath $base -PathType Container) {
            $finalPath = $base
            if ($expValue) {
                $cleanExp = ($expValue -replace '[^0-9\.]', '').Replace(".0","")
                # Improved exposure folder matching - exact match first, then prefix match
                $expDir = Get-ChildItem -LiteralPath $base -Directory -ErrorAction Stop | Where-Object {
                    $folderExp = ($_.Name -replace '[^0-9\.]', '').Replace(".0","")
                    $folderExp -eq $cleanExp 
                } | Select-Object -First 1
                
                # If no exact match found, try prefix match (for backwards compatibility)
                if (-not $expDir) {
                    $expDir = Get-ChildItem -LiteralPath $base -Directory -ErrorAction Stop | Where-Object {
                        $_.Name -like "$cleanExp*" 
                    } | Sort-Object Name | Select-Object -First 1
                }
                
                if ($expDir) { $finalPath = $expDir.FullName } else { continue }
            }
            $display = $finalPath.Replace($cb, "").TrimStart("\")
            return @{Path=$finalPath; Mode=$m; Display="$display"}
        }
    }
    return $null
}

function AngleDiff180($a, $b) {
    $diff = [Math]::Abs($a - $b) % 180
    if ($diff -gt 90) { $diff = 180 - $diff }
    return $diff
}

# SCANNER - iterates filter-group folders (L, RGB, Ha...) then date sub-folders
foreach ($fFolder in (Get-ChildItem -LiteralPath $lightsRoot -Directory -ErrorAction Stop)) {
    foreach ($dFolder in (Get-ChildItem -LiteralPath $fFolder.FullName -Directory -ErrorAction Stop)) {
        $sessionDate = ($dFolder.Name -split ' ')[0]
        # DEBUG: print the path and date folder name
        # Write-Host "[DEBUG] Processing session folder: $($dFolder.FullName) | date var: $sessionDate" -ForegroundColor DarkGray
        $fitsFiles = Get-ChildItem -LiteralPath $dFolder.FullName -Filter "*.fit*" -ErrorAction Stop
        if (!$fitsFiles) { continue }
        $sample = $fitsFiles | Select-Object -First 1
        $fName = $sample.Name
        $fFound = $null
        $foundIn = $null

        # 1. DETECT CAMERA FROM THIS SESSION'S FITS FILENAME
        $sessionCamFull = $camShort   # fallback to path-level short name
        if ($fName -match '_(?<camf>(?:ASI)?\d{3,4}M[MCPmcp]+)_') {
            $sessionCamFull = ("ASI" + $Matches['camf']) -replace '^ASIASI','ASI'
        }
        $sessionCamType  = if ($sessionCamFull -match "MM") { "Mono" } else { "OSC" }
        $sessionCalibRoot = Join-Path -Path $baseZ -ChildPath "Calibration"
        $sessionCalibBase = Join-Path -Path $sessionCalibRoot -ChildPath $sessionCamFull
        $camInFile = $sessionCamFull -replace '^ASI',''

        if (-not $camMap.ContainsKey($sessionCamFull)) {
            $camMap[$sessionCamFull] = @{ Observed = @(); Sanitized = @(); Targets = @(); Mappings = @() }
        }

        # 2. EXTRACT RAW FILTER FROM FILENAME
        $rawFilt = "None"  # Default to None if filter cannot be determined
        # Try to extract filter between camera and _gain pattern
        if ($fName -match "_$( [regex]::Escape($camInFile) )_(?<realFilt>[^_]+)?_gain") {
            $capturedFilt = $Matches['realFilt']
            if ($capturedFilt -and $capturedFilt.Trim() -ne "") {
                $rawFilt = $capturedFilt
            }
            # If nothing captured or empty, rawFilt remains "None"
        }
        # DEBUG: Uncomment next line to see filename parsing
        # Write-Host "[DEBUG] File: $fName | Cam: $camInFile | Filter: $rawFilt" -ForegroundColor DarkGray

        # 3. MAP TO NORMALIZED FILTER + WBPP TARGET
        $fm        = Get-FilterMap $rawFilt $sessionCamType
        $filt      = $fm.Filter
        $targetGrp = $fm.Target

        # Optional: Merge IRC/None into L for OSC cameras if user agrees
        if ($mergeFilters -and $sessionCamType -eq "OSC" -and ($rawFilt -eq "None" -or $rawFilt -eq "IRC" -or $rawFilt -eq "Trib")) {
            $filt = "L"
            $targetGrp = "RGB"  # Ensure unified target group for merged filters
        }

        $camMap[$sessionCamFull].Observed += $rawFilt
        $camMap[$sessionCamFull].Sanitized += $filt
        $camMap[$sessionCamFull].Targets += $targetGrp
        $camMap[$sessionCamFull].Mappings += ("$rawFilt -> $filt -> $targetGrp")

        # 3. METADATA PARSING (Angle fix)
        $valExp  = if ($fName -match "_(\d+\.?\d*)s_") { $Matches[1] } else { "300" }
        $curExp  = $valExp.Replace(".0","") + "s"
        $curGain = if ($fName -match "_gain(\d+)_") { $Matches[1] } else { "120" }
        $curTemp = if ($fName -match "_(-?\d+\.?\d*)C_") { $Matches[1] } else { "-20" }
        $curAngRaw = if ($fName -match "_(\d+)deg_") { $Matches[1] } else { $null }
        $numCurAng = if ($curAngRaw) { [int]$curAngRaw } else { -999 }
        $angDisp = if ($numCurAng -eq -999) { "Unknown" } else { "${numCurAng}deg" }

        # --- Add newline before session header for readability ---
        Write-Host ""
        Write-Host "  [$sessionCamFull/$sessionCamType] $sessionDate | Filter: $rawFilt -> $filt | Target: $targetGrp | Angle: $angDisp" -ForegroundColor Green

        $roundT = "$([Math]::Round([double]$curTemp / 5.0) * 5)C"
        $sTag = "Session_${sessionDate}_Filter_${filt}_Target_${targetGrp}_Gain_${curGain}_Exp_${curExp}_Cam_${sessionCamFull}"
        $cTag = "Gain_${curGain}_Temp_${roundT}_Exp_${curExp}_Cam_${sessionCamFull}"

        $pendingLinks += [PSCustomObject]@{ Type="Lights"; Tag=$sTag; Src=$dFolder.FullName; Display="Good\$($fFolder.Name)\$sessionDate"; Cam=$sessionCamFull }

        # --- FLATS SEARCH ---
        # Build list of all raw filter names that map to the same normalized filter,
        # so a flat shot as "IRC" is accepted for lights shot as "L" (and vice versa),
        # but "None" and "L" are kept separate (different flat corrections).
        $flatAliases = @()
        $flatAliases += $rawFilt
        # Remove addition of normalized filters to $flatAliases
        # $flatAliases += $filt
        # Add known synonyms for the raw filter only
        switch ($rawFilt) {
            "L"    { $flatAliases += "IRC"; $flatAliases += "Trib" }
            "H"    { $flatAliases += "Ha" }
            "Ha"   { $flatAliases += "H" }
            "S"    { $flatAliases += "SII" }
            "SII"  { $flatAliases += "S" }
            "O"    { $flatAliases += "OIII" }
            "OIII" { $flatAliases += "O" }
            "HO"   { $flatAliases += "UHC" }
            "UHC"  { $flatAliases += "HO" }
        }
        $flatAliases = $flatAliases | Select-Object -Unique
        
        
        # DEBUG: Show filter aliases
        # Write-Host "[DEBUG] Raw filter: $rawFilt | Aliases: $($flatAliases -join ', ')" -ForegroundColor Magenta

        # --- Robust filter match: look for filter as separate word or after date ---
        function FlatFolderMatchesFilter($folderName, $aliases) {
            if ([string]::IsNullOrWhiteSpace($folderName) -or $folderName.Length -lt 8) {
                return $false
            }

            $datePart = $folderName.Substring(0,8)
            $rest = $folderName.Substring(8).TrimStart()
            $filtPart = $null
            if ($rest -match '^(filt)?([^ _]+)') { $filtPart = $Matches[2] }
            
            # DEBUG: Uncomment for debugging filter matching
            # Write-Host "[DEBUG] FlatFolder: $folderName | Parsed date: $datePart | Parsed filter: $filtPart | Aliases: $($aliases -join ', ')" -ForegroundColor DarkGray
            
            foreach ($alias in $aliases) {
                if ($filtPart -eq $alias) { 
                    # Write-Host "[DEBUG]   MATCH: $folderName <-> $alias" -ForegroundColor Green
                    return $true 
                }
            }
            # Write-Host "[DEBUG]   SKIP: $folderName (no filter match)" -ForegroundColor Yellow
            return $false
        }

        # Known filters for validation
        $knownFilters = @('L','None','R','G','B','H','Ha','S','SII','O','OIII','IRC','Trib','HO','UHC','SO')

        # Helper: parse date in yy.MM.dd format only
        function ParseDate($dateStr) {
            try { 
                $parsedDate = [datetime]::ParseExact($dateStr, 'yy.MM.dd', $null)
                # Fix year interpretation for 2-digit years (assume 2000s for years 00-99)
                if ($parsedDate.Year -lt 2000) {
                    $parsedDate = $parsedDate.AddYears(100)
                }
                return $parsedDate 
            } catch { 
                return $null
            }
        }

        # Collect all candidate flats from Master and Source and always show menu for user selection
        $avail = @()
        # Write-Host "[DEBUG] Scanning flat folders in: $sessionCalibBase\*\flats\$telSetup" -ForegroundColor Cyan
        foreach($m in @("Master","Source")) {
            $p = Join-Path -Path $sessionCalibBase -ChildPath "$m\flats\$telSetup"
            if (Test-Path -LiteralPath $p -PathType Container) {
                Get-ChildItem -LiteralPath $p -Directory -ErrorAction Stop | ForEach-Object {
                    $item = $_ | Select-Object -Property Name, FullName
                    $folderName = $item.Name
                    $flatFilt = $null; $ang = $null; $warn = $false
                    if ($folderName.Length -lt 8) {
                        $warn = $true
                        Write-Host "[!] Flat folder '$folderName' is too short to contain a yy.MM.dd date prefix" -ForegroundColor Red
                    } else {
                        $date = $folderName.Substring(0,8)
                        $rest = $folderName.Substring(8).TrimStart()
                        if ($rest -match '^(filt)?([^ _]+)') { $flatFilt = $Matches[2] }
                        else { $warn = $true; Write-Host "[!] Flat folder '$folderName' does not match expected pattern (date filt<Filter> or date <Filter>)" -ForegroundColor Red }
                        # Validate date
                        $dateOk = $false
                        $d = ParseDate $date
                        if ($d) { $dateOk = $true } else { $dateOk = $false }
                        if (-not $dateOk) { $warn = $true; Write-Host "[!] Invalid date in flat folder: $folderName" -ForegroundColor Red }
                        # Validate filter
                        if ($flatFilt -and ($knownFilters -notcontains $flatFilt)) { $warn = $true; Write-Host "[!] Unknown filter '$flatFilt' in flat folder: $folderName" -ForegroundColor Red }
                        # Validate angle if present (optional)
                        if ($folderName -match '(\d+)deg') { $ang = $Matches[1] } else { $ang = $null }
                        if ($ang -and ($ang -notmatch '^\d+$')) { $warn = $true; Write-Host "[!] Invalid angle in flat folder: $folderName" -ForegroundColor Red }
                        # Write-Host "[DEBUG] Found flat folder: $folderName | Filter: $flatFilt | Date: $date | Angle: $ang" -ForegroundColor Magenta
                        $obj = [PSCustomObject]@{ Name = $item.Name; FullName = $item.FullName; Origin = $m }
                        $avail += $obj
                    }
                }
            }
        }

        if ($avail.Count -eq 0) {
            $lightsDisplay = "Good\$($fFolder.Name)\$sessionDate"
            Write-Host "`n[!] Missing Flats for $lightsDisplay (CalibRoot: $sessionCalibBase) (Original: $rawFilt, Target: $targetGrp, Angle: $angDisp)" -ForegroundColor Yellow
        } else {
            # Determine scores and default selection
            $defaultIdx = -1
            $fullGreenIdx = -1
            $shownAny = $false
            $masterGreenIdx = -1
            $shownIdx = @()
            for ($i=0; $i -lt $avail.Count; $i++) {
                $fDirName = $avail[$i].Name
                $notBad = ($fDirName -notmatch "\(")
                # date parse from folder name (dd.MM.yy or dd.MM.yyyy)
                $candDateStr = if ($fDirName -match '(?<!\d)(\d{2}\.\d{2}\.\d{2,4})(?!\d)') { $Matches[1] } else { $null }
                $candDate = $null
                if ($candDateStr) {
                    $candDate = ParseDate $candDateStr
                }
                # session date parse
                $sessDateObj = $null
                if ($sessionDate -match '(?<!\d)(\d{2}\.\d{2}\.\d{2,4})(?!\d)') {
                    $sd = $Matches[1]
                    $sessDateObj = ParseDate $sd
                }
                $dateDiff = 9999
                if ($candDate -and $sessDateObj) { $dateDiff = [Math]::Abs((New-TimeSpan -Start $candDate -End $sessDateObj).Days) }

                $fAng = if ($fDirName -match "(\d+)deg") { [int]$Matches[1] } else { -888 }
                if ($numCurAng -ne -999 -and $fAng -ne -888) {
                    $angDiff = AngleDiff180 $numCurAng $fAng
                } else {
                    $angDiff = 9999
                }
                $angOk = ($numCurAng -ne -999 -and $fAng -ne -888 -and $angDiff -le 2)
                $filtOk = FlatFolderMatchesFilter $fDirName $flatAliases

                $score = 0
                if ($filtOk) { $score += 10 }
                if ($angOk) { $score += 5 }
                if ($dateDiff -le 2) { $score += 3 }

                $avail[$i] | Add-Member -NotePropertyName Score -NotePropertyValue $score -Force
                $avail[$i] | Add-Member -NotePropertyName DateDiff -NotePropertyValue $dateDiff -Force
                $avail[$i] | Add-Member -NotePropertyName Ang -NotePropertyValue $fAng -Force
                $avail[$i] | Add-Member -NotePropertyName FiltOk -NotePropertyValue $filtOk -Force

                # Smart flat matching logic with priority system (fixed date vs angle priority)
                $dateMatch = ($dateDiff -eq 0)
                $angMatch = ($numCurAng -ne -999 -and $fAng -ne -888 -and $angDiff -le 2)
                
                # Calculate date relationship for smart matching
                $dateInPast = ($candDate -and $sessDateObj -and $candDate -lt $sessDateObj)
                $dateInNearFuture = ($candDate -and $sessDateObj -and $candDate -gt $sessDateObj -and $dateDiff -le 2)
                $dateNear = ($dateDiff -le 2)
                
                $rowColor = 'DarkGray'
                if ($filtOk) {
                    $shownIdx += $i
                    if ($dateMatch -and $angMatch) {
                        # Perfect match: exact date + angle
                        $rowColor = 'Green'
                        if ($fullGreenIdx -eq -1) { $fullGreenIdx = $i }
                        if ($avail[$i].Origin -eq 'Master' -and $masterGreenIdx -eq -1) { $masterGreenIdx = $i }
                    } elseif ($dateMatch -or $dateNear) {
                        # Good match by DATE: exact or close date (priority over angles)
                        $rowColor = 'Yellow'
                    } elseif ($angMatch) {
                        # Match by ANGLE only (lower priority)
                        $rowColor = 'Yellow'
                    }
                    # Note: Blue coloring for smart matches will be determined later in the display loop
                }
            }
            # Default selection priority with the corrected logic:
            # 1. Master among green matches (exact date+angle match)
            # 2. Any green match (exact date+angle match)
            # 3. Master among yellow DATE matches (date has priority over angle)
            # 4. Any yellow DATE match
            # 5. Master among blue matches: closest past date
            # 6. Any blue match: closest past date
            # 7. Master among blue matches: closest future date (1-2 days)
            # 8. Any blue match: closest future date (1-2 days)
            # 9. Master among yellow ANGLE matches (angles have lowest priority)
            # 10. Any yellow ANGLE match
            $masterYellowDateIdx = -1
            $yellowDateIdx = -1
            $masterYellowAngleIdx = -1
            $yellowAngleIdx = -1
            $masterClosestPastIdx = -1
            $closestPastIdx = -1
            $masterClosestFutureIdx = -1
            $closestFutureIdx = -1
            
            $closestPastDays = 9999
            $closestFutureDays = 9999
            
            # Collect priority indices while searching for the nearest dates
            for ($i=0; $i -lt $avail.Count; $i++) {
                if (-not $avail[$i].FiltOk) { continue }
                
                $fDirName = $avail[$i].Name
                $candDateStr = if ($fDirName -match '(?<!\d)(\d{2}\.\d{2}\.\d{2,4})(?!\d)') { $Matches[1] } else { $null }
                $candDate = $null
                if ($candDateStr) {
                    $candDate = ParseDate $candDateStr
                }
                
                $dateDiff = 9999
                if ($candDate -and $sessDateObj) { $dateDiff = [Math]::Abs((New-TimeSpan -Start $candDate -End $sessDateObj).Days) }
                
                $dateMatch = ($dateDiff -eq 0)
                $dateNear = ($dateDiff -le 2)
                $fAng = if ($fDirName -match "(\d+)deg") { [int]$Matches[1] } else { -888 }
                $angMatch = ($numCurAng -ne -999 -and $fAng -ne -888 -and (AngleDiff180 $numCurAng $fAng) -le 2)
                
                $dateInPast = ($candDate -and $sessDateObj -and $candDate -lt $sessDateObj)
                $dateInNearFuture = ($candDate -and $sessDateObj -and $candDate -gt $sessDateObj -and $dateDiff -le 2)
                
                # Classify match types with date-first priority
                if ($dateMatch -and $angMatch) {
                    # Green matches were already handled in the first pass
                } elseif ($dateMatch -or $dateNear) {
                    # Yellow DATE matches have the highest priority among yellow matches
                    if ($yellowDateIdx -eq -1) { $yellowDateIdx = $i }
                    if ($avail[$i].Origin -eq 'Master' -and $masterYellowDateIdx -eq -1) { $masterYellowDateIdx = $i }
                } elseif ($dateInPast) {
                    # Blue past matches: prefer the nearest/newest date and filter angles less strictly
                    $pastAngleConflict = ($numCurAng -ne -999 -and $fAng -ne -888 -and (AngleDiff180 $numCurAng $fAng) -gt 30)
                    if (-not $pastAngleConflict -and $dateDiff -lt $closestPastDays) {
                        $closestPastDays = $dateDiff
                        $closestPastIdx = $i
                        if ($avail[$i].Origin -eq 'Master') { $masterClosestPastIdx = $i }
                    } elseif (-not $pastAngleConflict -and $dateDiff -eq $closestPastDays) {
                        # If the date is the same, prefer a matching angle
                        if ($angMatch -and -not ($numCurAng -ne -999 -and $avail[$closestPastIdx].Ang -ne -888 -and (AngleDiff180 $numCurAng $avail[$closestPastIdx].Ang) -le 2)) {
                            $closestPastIdx = $i
                        }
                        if ($avail[$i].Origin -eq 'Master') {
                            if ($masterClosestPastIdx -eq -1) {
                                $masterClosestPastIdx = $i
                            } elseif ($angMatch -and -not ($numCurAng -ne -999 -and $avail[$masterClosestPastIdx].Ang -ne -888 -and (AngleDiff180 $numCurAng $avail[$masterClosestPastIdx].Ang) -le 2)) {
                                $masterClosestPastIdx = $i
                            }
                        }
                    }
                } elseif ($dateInNearFuture) {
                    # Blue near-future matches: prefer the nearest date, but reject incompatible angles
                    $futureAngleConflict = ($numCurAng -ne -999 -and $fAng -ne -888 -and (AngleDiff180 $numCurAng $fAng) -gt 10)
                    if (-not $futureAngleConflict -and $dateDiff -lt $closestFutureDays) {
                        $closestFutureDays = $dateDiff
                        $closestFutureIdx = $i
                        if ($avail[$i].Origin -eq 'Master') { $masterClosestFutureIdx = $i }
                    } elseif (-not $futureAngleConflict -and $dateDiff -eq $closestFutureDays) {
                        # If the date is the same, prefer a matching angle
                        if ($angMatch -and -not ($numCurAng -ne -999 -and $avail[$closestFutureIdx].Ang -ne -888 -and (AngleDiff180 $numCurAng $avail[$closestFutureIdx].Ang) -le 2)) {
                            $closestFutureIdx = $i
                        }
                        if ($avail[$i].Origin -eq 'Master') {
                            if ($masterClosestFutureIdx -eq -1) {
                                $masterClosestFutureIdx = $i
                            } elseif ($angMatch -and -not ($numCurAng -ne -999 -and $avail[$masterClosestFutureIdx].Ang -ne -888 -and (AngleDiff180 $numCurAng $avail[$masterClosestFutureIdx].Ang) -le 2)) {
                                $masterClosestFutureIdx = $i
                            }
                        }
                    }
                } elseif ($angMatch) {
                    # Yellow ANGLE matches have the lowest priority among yellow matches
                    if ($yellowAngleIdx -eq -1) { $yellowAngleIdx = $i }
                    if ($avail[$i].Origin -eq 'Master' -and $masterYellowAngleIdx -eq -1) { $masterYellowAngleIdx = $i }
                }
            }
            
            # Pick the default index: DATE is the main criterion, ANGLE refines it
            if ($masterGreenIdx -ge 0) { $defaultIdx = $masterGreenIdx }
            elseif ($fullGreenIdx -ge 0) { $defaultIdx = $fullGreenIdx }
            elseif ($masterYellowDateIdx -ge 0) { $defaultIdx = $masterYellowDateIdx }
            elseif ($yellowDateIdx -ge 0) { $defaultIdx = $yellowDateIdx }
            elseif ($masterClosestPastIdx -ge 0) { $defaultIdx = $masterClosestPastIdx }
            elseif ($closestPastIdx -ge 0) { $defaultIdx = $closestPastIdx }
            elseif ($masterClosestFutureIdx -ge 0) { $defaultIdx = $masterClosestFutureIdx }
            elseif ($closestFutureIdx -ge 0) { $defaultIdx = $closestFutureIdx }
            elseif ($masterYellowAngleIdx -ge 0) { $defaultIdx = $masterYellowAngleIdx }
            elseif ($yellowAngleIdx -ge 0) { $defaultIdx = $yellowAngleIdx }
            else { $defaultIdx = -1 }

            # Show both equivalent angles as a hint
            Write-Host "Select flats from calibration root: $sessionCalibBase" -ForegroundColor Cyan
            Write-Host "Session: Good\$($fFolder.Name)\$sessionDate" -ForegroundColor Cyan
            if ($numCurAng -ne -999) {
                $altAng = ($numCurAng + 180) % 360
                Write-Host "Angle: ${numCurAng}deg or ${altAng}deg" -ForegroundColor Cyan
            }
            else {
                Write-Host "Angle: Unknown" -ForegroundColor Cyan
            }

            # Create display mapping for shown items with continuous indices
            $displayItems = @()
            $originalToDisplayMap = @{}
            $displayToOriginalMap = @{}
            $displayDefaultIdx = -1
            
            for ($i=0; $i -lt $avail.Count; $i++) {
                $ent = $avail[$i]
                if (-not $ent.FiltOk) { continue }
                
                $displayIdx = $displayItems.Count
                $displayItems += $ent
                $originalToDisplayMap[$i] = $displayIdx
                $displayToOriginalMap[$displayIdx] = $i
                
                # Map default index from original to display
                if ($defaultIdx -eq $i) { $displayDefaultIdx = $displayIdx }
            }

            if ($displayItems.Count -eq 0) {
                $lightsDisplay = "Good\$($fFolder.Name)\$sessionDate"
                Write-Host "`n[!] No matching Flats for $lightsDisplay (CalibRoot: $sessionCalibBase) (Original: $rawFilt, Target: $targetGrp, Angle: $angDisp)" -ForegroundColor Yellow
            } else {
            $shownAny = $false
            for ($displayIdx=0; $displayIdx -lt $displayItems.Count; $displayIdx++) {
                $ent = $displayItems[$displayIdx]
                $originalIdx = $displayToOriginalMap[$displayIdx]
                $shownAny = $true

                # Date and angle colors with the improved matching logic
                $dateColor = 'White'; $angColor = 'White'; $rowColor = 'DarkGray'
                $dateMatch = $ent.DateDiff -eq 0
                $dateNear = ($ent.DateDiff -le 2)
                
                # Recalculate dates for the current row
                $candDateStr = if ($ent.Name -match '(?<!\d)(\d{2}\.\d{2}\.\d{2,4})(?!\d)') { $Matches[1] } else { $null }
                $candDate = $null
                if ($candDateStr) {
                    $candDate = ParseDate $candDateStr
                }
                
                $dateInPast = ($candDate -and $sessDateObj -and $candDate -lt $sessDateObj)
                $dateInNearFuture = ($candDate -and $sessDateObj -and $candDate -gt $sessDateObj -and $ent.DateDiff -le 2)

                if ($numCurAng -ne -999 -and $ent.Ang -ne -888) {
                    $angDiff = AngleDiff180 $numCurAng $ent.Ang
                } else {
                    $angDiff = 9999
                }

                $angMatch = ($numCurAng -ne -999 -and $ent.Ang -ne -888 -and $angDiff -le 2)
                $angNear = ($numCurAng -ne -999 -and $ent.Ang -ne -888 -and $angDiff -le 5)
                
                # Determine whether this flat is one of the selected nearest matches (using original indices)
                $isClosestPast = ($closestPastIdx -eq $originalIdx -or $masterClosestPastIdx -eq $originalIdx)
                $isClosestFuture = ($closestFutureIdx -eq $originalIdx -or $masterClosestFutureIdx -eq $originalIdx)
                
                # Color coding for dates
                if ($dateMatch) { $dateColor = 'Green' } 
                elseif ($dateNear) { $dateColor = 'Yellow' }
                elseif ($isClosestPast -or $isClosestFuture) { $dateColor = 'Blue' }
                
                # Color coding for angles
                if ($angMatch) { $angColor = 'Green' } 
                elseif ($angNear) { $angColor = 'Yellow' }

                # Determine the whole-row color with the new logic
                if ($dateMatch -and $angMatch) { $rowColor = 'Green' }
                elseif ($dateMatch -or $angMatch) { $rowColor = 'Yellow' }
                elseif ($isClosestPast -or $isClosestFuture) { $rowColor = 'Blue' }
                elseif ($dateNear -or $angNear) { $rowColor = 'Yellow' }
                else { $rowColor = 'DarkGray' }
                $relPath = $ent.FullName.Replace($sessionCalibBase, '').TrimStart('\')

                # Build a row with colored date and angle values (using display index)
                $idxStr = "[$displayIdx]".PadRight(5)
                $origStr = "[$($ent.Origin)]".PadRight(10)
                $nameParts = $ent.Name -split ' '
                $dateStr = $nameParts[0]
                $restStr = $ent.Name.Substring($dateStr.Length).TrimStart()
                if ($ent.Ang -ne -888) { $angStr = "$($ent.Ang)deg" } else { $angStr = "" }

                # Remove the angle from restStr when present
                if ($angStr -ne "") {
                    $restStr = $restStr -replace "\s*\b$($ent.Ang)deg\b", ""
                }
                # Align the arrow at column 53 when possible
                $leftPart = " $idxStr $origStr $dateStr $restStr $angStr"
                $arrowCol = 53
                $arrowPad = ""
                if ($leftPart.Length -lt $arrowCol) { 
                    $arrowPad = ' ' * ($arrowCol - $leftPart.Length) 
                } else {
                    # If the row exceeds the limit, add a minimal gap
                    $arrowPad = "  "
                }
                $arrowPad += "<--"

                # Use green for the default selection index (using display index)
                $idxColor = if ($displayIdx -eq $displayDefaultIdx) { 'Green' } else { 'White' }

                # Print the row with colored date/angle values and Blue support for smart flat matches
                if ($rowColor -eq 'Green') {
                    Write-Host " " -NoNewline
                    Write-Host $idxStr -NoNewline -ForegroundColor $idxColor
                    Write-Host " $origStr " -NoNewline -ForegroundColor White
                    Write-Host $dateStr -NoNewline -ForegroundColor Green
                    Write-Host (" $restStr ") -NoNewline -ForegroundColor White
                    if ($angStr -ne "") { Write-Host $angStr -NoNewline -ForegroundColor Green }
                    Write-Host ("$arrowPad $relPath") -ForegroundColor Green
                } elseif ($rowColor -eq 'Yellow') {
                    Write-Host " " -NoNewline
                    Write-Host $idxStr -NoNewline -ForegroundColor $idxColor
                    Write-Host " $origStr " -NoNewline -ForegroundColor White
                    Write-Host $dateStr -NoNewline -ForegroundColor $dateColor
                    Write-Host (" $restStr ") -NoNewline -ForegroundColor White
                    if ($angStr -ne "") { Write-Host $angStr -NoNewline -ForegroundColor $angColor }
                    Write-Host ("$arrowPad $relPath") -ForegroundColor Yellow
                } elseif ($rowColor -eq 'Blue') {
                    # Smart match: closest past or near future flats
                    Write-Host " " -NoNewline
                    Write-Host $idxStr -NoNewline -ForegroundColor $idxColor
                    Write-Host " $origStr " -NoNewline -ForegroundColor White
                    Write-Host $dateStr -NoNewline -ForegroundColor Blue
                    Write-Host (" $restStr ") -NoNewline -ForegroundColor White
                    if ($angStr -ne "") { Write-Host $angStr -NoNewline -ForegroundColor $angColor }
                    Write-Host ("$arrowPad $relPath") -ForegroundColor Blue
                } else {
                    # Gray rows also use componentized output for proper alignment
                    Write-Host " " -NoNewline
                    Write-Host $idxStr -NoNewline -ForegroundColor $idxColor
                    Write-Host " $origStr $dateStr $restStr $angStr$arrowPad $relPath" -ForegroundColor DarkGray
                }
            }

            # Default index: use the previously computed fullGreenIdx or masterGreenIdx
            # If there are no exact matches, use a date-matching filter as the default
            if ($displayDefaultIdx -eq -1) {
                for ($displayIdx=0; $displayIdx -lt $displayItems.Count; $displayIdx++) {
                    $ent = $displayItems[$displayIdx]
                    if ($ent.FiltOk -and $ent.DateDiff -eq 0) {
                        $displayDefaultIdx = $displayIdx
                        break
                    }
                }
            }
            
            # If there is still no default, use the first available item
            if ($displayDefaultIdx -eq -1 -and $displayItems.Count -gt 0) {
                $displayDefaultIdx = 0
            }
            
            if ($displayDefaultIdx -ge 0 -and $displayDefaultIdx -lt $displayItems.Count) { 
                $dtext = " $displayDefaultIdx" 
            } else { 
                $dtext = " 0"  # fallback
                $displayDefaultIdx = 0
            }
            
            # Input validation loop
            do {
                $prompt = "Select Index (Enter to accept default$dtext)"
                $idx = Read-Host -Prompt $prompt
                $validInput = $false
                $selectedIdx = -1
                
                if ($idx.Trim() -eq "") {
                    # Default selection
                    $selectedIdx = $displayDefaultIdx
                    $validInput = $true
                } elseif ($idx -match '^\d+$') {
                    $selectedIdx = [int]$idx
                    if ($selectedIdx -ge 0 -and $selectedIdx -lt $displayItems.Count) {
                        $validInput = $true
                    } else {
                        Write-Host "[!] Invalid index. Please enter a number between 0 and $($displayItems.Count - 1)" -ForegroundColor Red
                    }
                } else {
                    Write-Host "[!] Invalid input. Please enter a number or press Enter for default" -ForegroundColor Red
                }
            } while (-not $validInput)
            
            Write-Host ""
            $fFound = $displayItems[$selectedIdx]
            $foundIn = $fFound.Origin
            }
        }

        if ($fFound) {
            # Clean leading slash for flats display too
            $fDisp = "$foundIn\flats\$telSetup\$($fFound.Name)"
            $pendingLinks += [PSCustomObject]@{ Type="Flats"; Tag=$sTag; Src=$fFound.FullName; Display=$fDisp; Cam=$sessionCamFull }

            # Check for potential keyword conflicts in flat files
            $flatFiles = Get-ChildItem -LiteralPath $fFound.FullName -Filter "*.fit*" -ErrorAction SilentlyContinue
            if ($flatFiles) {
                foreach ($flatFile in $flatFiles) {
                    $fileName = $flatFile.Name
                    # Check if filename contains FILTER keyword (case-insensitive) that conflicts with our mapping
                    if ($fileName -match "(?i)filter[_-]([^_\s\.]+)") {
                        $fileFilter = $Matches[1]
                        if ($fileFilter -and $fileFilter -ne $filt) {
                            Write-Host "[!] KEYWORD CONFLICT WARNING: File '$fileName'" -ForegroundColor Red
                            Write-Host "    File contains FILTER=$fileFilter but our mapping expects FILTER=$filt" -ForegroundColor Red
                            Write-Host "    WBPP may use the filename keyword instead of our normalized value" -ForegroundColor Yellow
                            Write-Host "    Consider renaming the file to use FILTER=$filt for consistency" -ForegroundColor Yellow
                        }
                    }
                }
            }

            # Flat-Darks
            $fSample = Get-ChildItem -LiteralPath $fFound.FullName -Filter "*.fit*" -ErrorAction Stop | Select-Object -First 1
            if ($fSample -and ($fSample.Name -match "_(\d+\.?\d*)m?s_")) {
                $fdExp = $Matches[1] + "s"
                $fd = Get-CalibPath "FlatDarks" $curGain $curTemp $fdExp $sessionCalibBase
                if ($fd) {
                $fdTag = "Exp_${fdExp}_Gain_${curGain}_Target_${targetGrp}_Filter_${filt}_Cam_${sessionCamFull}"
                    $pendingLinks += [PSCustomObject]@{ Type="FlatDarks"; Tag=$fdTag; Src=$fd.Path; Display=$fd.Display; Cam=$sessionCamFull }
                    # Check for potential keyword conflicts in flat-dark files
                    $flatDarkFiles = Get-ChildItem -LiteralPath $fd.Path -Filter "*.fit*" -ErrorAction SilentlyContinue
                    if ($flatDarkFiles) {
                        foreach ($flatDarkFile in ($flatDarkFiles | Select-Object -First 3)) {  # Check only first 3 files for performance
                            $fileName = $flatDarkFile.Name
                            if ($fileName -match "(?i)filter[_-]([^_\s\.]+)") {
                                $fileFilter = $Matches[1]
                                if ($fileFilter -and $fileFilter -ne $filt) {
                                    Write-Host "[!] KEYWORD CONFLICT WARNING: FlatDark file '$fileName'" -ForegroundColor Red
                                    Write-Host "    File contains FILTER=$fileFilter but our mapping expects FILTER=$filt" -ForegroundColor Red
                                }
                            }
                        }
                    }
                }
            }
        }

        # --- DARKS & BIAS ---
        # Add Target and Filter to calibration frame symlink names
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
    # --- Calculate total size for copy operation ---
    $uniqueSourcePaths = $pendingLinks.Src | Select-Object -Unique
    $totalSize = 0
    foreach ($srcPath in $uniqueSourcePaths) {
        if (Test-Path -LiteralPath $srcPath) {
            $size = (Get-ChildItem -LiteralPath $srcPath -Recurse -File -ErrorAction Stop | Measure-Object -Property Length -Sum).Sum
            if ($null -ne $size) {
                $totalSize += $size
            }
        }
    }
    $totalSizeGB = [Math]::Round($totalSize / 1GB, 2)

    # --- Ask user for action ---
    Write-Host "`nCreate project structure in '$sourcePath'" -ForegroundColor Cyan
    Write-Host " [1] Create symlinks (recommended, instant)" -ForegroundColor White
    Write-Host " [2] Copy all source files ($totalSizeGB GB required)" -ForegroundColor White
    
    $choice = Read-Host "Select option (default is 1)"
    if ($choice -eq '' -or $choice -eq '1') {
        # --- ACTION 1: Create Symlinks ---
        Write-Host "`nCreating symlinks..." -ForegroundColor Green
        foreach ($l in $pendingLinks) {
            $targetTypePath = Join-Path -Path $sourcePath -ChildPath $l.Type
            $target = Join-Path -Path $targetTypePath -ChildPath $l.Tag
            New-CreateProjectSymbolicLink -Path $target -Target $l.Src
        }
        Write-Host "`n[DONE] Symlinks created. Project Root: $setupRoot" -ForegroundColor Yellow

    } elseif ($choice -eq '2') {
        # --- ACTION 2: Copy Files using Robocopy ---
        Write-Host "`nCopying files... This may take a while." -ForegroundColor Green
        $uniqueLinks = $pendingLinks | Sort-Object -Property Src -Unique

        foreach ($l in $uniqueLinks) {
            # We need to find all target tags that point to this unique source
            $targetsForSrc = $pendingLinks | Where-Object { $_.Src -eq $l.Src }
            
            foreach ($targetEntry in $targetsForSrc) {
                $targetTypePath = Join-Path -Path $sourcePath -ChildPath $targetEntry.Type
                $targetPath = Join-Path -Path $targetTypePath -ChildPath $targetEntry.Tag
                Write-Host "Copying from $($l.Src) to $targetPath" -ForegroundColor DarkGray
                Copy-CreateProjectDirectory -Source $l.Src -Destination $targetPath
            }
        }
        Write-Host "`n[DONE] Files copied. Project Root: $setupRoot" -ForegroundColor Yellow
    } else {
        Write-Host "`n[ABORTED] No action taken." -ForegroundColor Red
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

Write-Host "`n--- IMPORTANT: Keyword Conflicts ---" -ForegroundColor Cyan
Write-Host "If you see KEYWORD CONFLICT warnings above:" -ForegroundColor Gray  
Write-Host "- WBPP may prioritize filename keywords over your custom FILTER keywords" -ForegroundColor Yellow
Write-Host "- This can cause calibration frame mismatches (e.g., Ha flats not matching H lights)" -ForegroundColor Yellow
Write-Host "- Consider renaming conflicting files to use normalized filter names (Ha->H, OIII->O, SII->S)" -ForegroundColor Yellow
Write-Host "- Consider removing conflicting keywords from calibration file (FILTER, GAIN etc)" -ForegroundColor Yellow
Write-Host "- Alternative: Use WBPP's 'Smart naming override' option and clear/reload file list" -ForegroundColor Gray

# --- After scan completion and camMap construction ---

$meta = @{
    Object = $astroObj
    Season = $season
    Scope = $telSetup
    GoodPath = $inputPath
    PixPath = $pixPath
    Cameras = @()
}
foreach ($cam in $camMap.Keys) {
    $entry = $camMap[$cam]
    $filters = ($entry.Sanitized | Select-Object -Unique)
    $targets = ($entry.Targets | Select-Object -Unique)
    # Calibration folders: look for Master roots for each camera
    $calibrationRoot = Join-Path -Path $baseZ -ChildPath "Calibration"
    $calibBase = Join-Path -Path $calibrationRoot -ChildPath $cam
    $calibFolders = @{
        Darks = Join-Path $calibBase "Master\darks"
        Biases = Join-Path $calibBase "Master\biases"
        Flats = Join-Path $calibBase "Master\flats"
        FlatDarks = Join-Path $calibBase "Master\flat-darks"
    }
    $meta.Cameras += @{
        Name = $cam
        Filters = $filters
        Targets = $targets
        CalibrationFolders = $calibFolders
    }
}

# Save JSON next to Pix/Good
$metaPath = Join-Path (Split-Path $pixPath -Parent) "project_meta.json"
if ($PSCmdlet.ShouldProcess($metaPath, "Write project metadata")) {
    $meta | ConvertTo-Json -Depth 6 | Out-File -Encoding UTF8 -LiteralPath $metaPath
    Write-Host "`n[INFO] Project metadata saved to: $metaPath" -ForegroundColor Cyan
}
