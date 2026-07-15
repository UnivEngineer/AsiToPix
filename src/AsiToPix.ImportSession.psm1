Set-StrictMode -Version Latest

function Read-AsiToPixRequiredValue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Prompt,

        [string]$DefaultValue = ""
    )

    do {
        if ([string]::IsNullOrWhiteSpace($DefaultValue)) {
            $value = (Read-Host $Prompt).Trim()
        } else {
            $value = (Read-Host "$Prompt [$DefaultValue]").Trim()
            if ([string]::IsNullOrWhiteSpace($value)) {
                $value = $DefaultValue
            }
        }

        if (-not [string]::IsNullOrWhiteSpace($value)) {
            return $value
        }

        Write-Host "[!] Value cannot be empty." -ForegroundColor Red
    } while ($true)
}

function Read-AsiToPixConfirmation {
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

function ConvertTo-AsiToPixNameToken {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $normalized = $Name.ToLowerInvariant()
    $normalized = $normalized -replace '\b(nebula|galaxy|cluster|region|the)\b', ' '
    $normalized = $normalized -replace '[^a-z0-9]+', ' '

    return @($normalized.Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries) | Select-Object -Unique)
}

function Get-AsiToPixCatalogIdentifier {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if ($Name -match '(?i)\b(?<catalog>M|NGC|IC|SH2|SH|LBN|LDN)\s*[- ]?\s*(?<number>\d+[A-Z]?)\b') {
        return "$($Matches["catalog"].ToUpperInvariant()) $($Matches["number"].ToUpperInvariant())"
    }

    return ""
}

function Get-AsiToPixNameMatch {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DetectedName,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]]$Candidates
    )

    if ([string]::IsNullOrWhiteSpace($DetectedName) -or $Candidates.Count -eq 0) {
        return @()
    }

    $detected = $DetectedName.Trim()
    $detectedLower = $detected.ToLowerInvariant()
    $detectedCompact = $detectedLower -replace '[^a-z0-9]+', ''
    $detectedTokens = @(ConvertTo-AsiToPixNameToken -Name $detected)
    $detectedCatalogIdentifier = Get-AsiToPixCatalogIdentifier -Name $detected

    $nameMatches = foreach ($candidate in $Candidates) {
        if ([string]::IsNullOrWhiteSpace($candidate)) { continue }

        $candidateLower = $candidate.ToLowerInvariant()
        $candidateCompact = $candidateLower -replace '[^a-z0-9]+', ''
        $candidateTokens = @(ConvertTo-AsiToPixNameToken -Name $candidate)
        $candidateCatalogIdentifier = Get-AsiToPixCatalogIdentifier -Name $candidate
        $score = 0

        if (-not [string]::IsNullOrWhiteSpace($detectedCatalogIdentifier)) {
            if ($candidateCatalogIdentifier -ne $detectedCatalogIdentifier) {
                continue
            }

            $score += 500
        }

        if ($candidateLower -eq $detectedLower) {
            $score += 1000
        }

        if ($candidateLower.Contains($detectedLower) -or $detectedLower.Contains($candidateLower)) {
            $score += 200
        }

        if (-not [string]::IsNullOrWhiteSpace($detectedCompact) -and
            ($candidateCompact.Contains($detectedCompact) -or $detectedCompact.Contains($candidateCompact))) {
            $score += 160
        }

        $sharedTokens = @($detectedTokens | Where-Object { $candidateTokens -contains $_ })
        if ($sharedTokens.Count -gt 0) {
            $score += 40 * $sharedTokens.Count
            if ($sharedTokens.Count -eq $detectedTokens.Count) {
                $score += 80
            }
        }

        if ($score -gt 0) {
            [PSCustomObject]@{
                Name  = $candidate
                Score = $score
            }
        }
    }

    return @($nameMatches | Sort-Object -Property @{ Expression = "Score"; Descending = $true }, Name)
}

function Read-AsiToPixNameSelection {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Kind,

        [Parameter(Mandatory = $true)]
        [string]$DetectedName,

        [string[]]$Candidates = @()
    )

    $nameMatches = @(Get-AsiToPixNameMatch -DetectedName $DetectedName -Candidates $Candidates)

    if ($nameMatches.Count -eq 1) {
        $candidate = $nameMatches[0].Name
        if (Read-AsiToPixConfirmation -Prompt "Use existing $Kind '$candidate' for detected '$DetectedName'?") {
            return [PSCustomObject]@{
                Name  = $candidate
                IsNew = $false
            }
        }

        $newName = Read-AsiToPixRequiredValue -Prompt "Enter destination $Kind name" -DefaultValue $DetectedName
        return [PSCustomObject]@{
            Name  = $newName
            IsNew = ($Candidates -notcontains $newName)
        }
    }

    if ($nameMatches.Count -gt 1) {
        Write-Host "`nMatching existing $Kind names for '${DetectedName}':" -ForegroundColor Cyan
        for ($i = 0; $i -lt $nameMatches.Count; $i++) {
            Write-Host " [$($i + 1)] $($nameMatches[$i].Name)" -ForegroundColor White
        }
        Write-Host " [0] Use a new name" -ForegroundColor White

        do {
            $answer = (Read-Host "Select $Kind index, press Enter for 1, or type a new name").Trim()
            if ([string]::IsNullOrWhiteSpace($answer)) {
                return [PSCustomObject]@{
                    Name  = $nameMatches[0].Name
                    IsNew = $false
                }
            }

            $selectedIndex = -1
            if ([int]::TryParse($answer, [ref]$selectedIndex)) {
                if ($selectedIndex -eq 0) {
                    $newName = Read-AsiToPixRequiredValue -Prompt "Enter destination $Kind name" -DefaultValue $DetectedName
                    return [PSCustomObject]@{
                        Name  = $newName
                        IsNew = ($Candidates -notcontains $newName)
                    }
                }

                if ($selectedIndex -ge 1 -and $selectedIndex -le $nameMatches.Count) {
                    return [PSCustomObject]@{
                        Name  = $nameMatches[$selectedIndex - 1].Name
                        IsNew = $false
                    }
                }
            } elseif (-not [string]::IsNullOrWhiteSpace($answer)) {
                return [PSCustomObject]@{
                    Name  = $answer
                    IsNew = ($Candidates -notcontains $answer)
                }
            }

            Write-Host "[!] Invalid selection." -ForegroundColor Red
        } while ($true)
    }

    Write-Host "[INFO] No matching existing $Kind found for '$DetectedName'." -ForegroundColor Yellow
    $value = Read-AsiToPixRequiredValue -Prompt "Enter destination $Kind name" -DefaultValue $DetectedName

    return [PSCustomObject]@{
        Name  = $value
        IsNew = ($Candidates -notcontains $value)
    }
}

function Get-AsiToPixLightFileInfo {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileName
    )

    $nameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    if ($nameWithoutExtension.EndsWith(".fit", [System.StringComparison]::OrdinalIgnoreCase)) {
        $nameWithoutExtension = [System.IO.Path]::GetFileNameWithoutExtension($nameWithoutExtension)
    }

    $objectName = $null
    $exposureSeconds = $null
    if ($nameWithoutExtension -match '^Light_(?<object>.+?)_(?<exp>\d+(?:\.\d+)?)s(?:_|$)') {
        $objectName = ($Matches["object"] -replace '_', ' ').Trim()
        $exposureSeconds = $Matches["exp"]
    }

    $cameraName = $null
    if ($nameWithoutExtension -match '_(?<camera>(?:ASI)?\d{3,4}M[MCP]?)_') {
        $cameraName = $Matches["camera"]
        if ($cameraName -notmatch '^ASI') {
            $cameraName = "ASI$cameraName"
        }
    }

    $filterName = "None"
    if ($cameraName) {
        $cameraToken = $cameraName -replace '^ASI', ''
        if ($nameWithoutExtension -match "_$([regex]::Escape($cameraToken))_(?<filter>[^_]*)_gain") {
            if (-not [string]::IsNullOrWhiteSpace($Matches["filter"])) {
                $filterName = $Matches["filter"].Trim()
            }
        } elseif ($nameWithoutExtension -match "_$([regex]::Escape($cameraName))_(?<filter>[^_]*)_gain") {
            if (-not [string]::IsNullOrWhiteSpace($Matches["filter"])) {
                $filterName = $Matches["filter"].Trim()
            }
        }
    }

    $capturedAt = $null
    if ($nameWithoutExtension -match '_(?<stamp>\d{8}-\d{6})(?:_|$)') {
        $capturedAt = [datetime]::ParseExact(
            $Matches["stamp"],
            "yyyyMMdd-HHmmss",
            [System.Globalization.CultureInfo]::InvariantCulture
        )
    }

    $telescopeName = $null
    if ($nameWithoutExtension -match '_-?\d+(?:\.\d+)?C_(?<scope>[^_]+)_\d+$') {
        $telescopeName = $Matches["scope"].Trim()
    }

    return [PSCustomObject]@{
        ObjectName      = $objectName
        ExposureSeconds = $exposureSeconds
        CameraName      = $cameraName
        FilterName      = $filterName
        CapturedAt      = $capturedAt
        TelescopeName   = $telescopeName
    }
}

function Get-AsiToPixNightDate {
    param(
        [Parameter(Mandatory = $true)]
        [datetime]$CapturedAt
    )

    $nightStart = $CapturedAt.Date
    if ($CapturedAt.Hour -lt 12) {
        $nightStart = $nightStart.AddDays(-1)
    }

    return $nightStart.ToString("yy.MM.dd", [System.Globalization.CultureInfo]::InvariantCulture)
}

function Get-AsiToPixSetupInfo {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SetupName
    )

    $scope = $SetupName
    $camera = $null
    if ($SetupName -match '^(?<scope>.+?) @ (?<camera>ASI\d+.*)$') {
        $scope = $Matches["scope"].Trim()
        $camera = $Matches["camera"].Trim()
    }

    return [PSCustomObject]@{
        Scope  = $scope
        Camera = $camera
    }
}

function Get-AsiToPixDetectedTelescope {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,

        [string]$FallbackTelescope = ""
    )

    $directory = Get-Item -LiteralPath $SourcePath
    if (-not $directory.PSIsContainer) {
        $directory = Get-Item -LiteralPath (Split-Path -Path $directory.FullName -Parent)
    }

    $current = $directory
    while ($null -ne $current) {
        if ($current.Name -ieq "Light" -and $null -ne $current.Parent) {
            return $current.Parent.Name
        }

        if ($null -ne $current.Parent -and $current.Parent.Name -ieq "Light") {
            if ($null -ne $current.Parent.Parent) {
                return $current.Parent.Parent.Name
            }
        }

        $current = $current.Parent
    }

    return $FallbackTelescope
}

function Get-AsiToPixDetectedObject {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath
    )

    $directory = Get-Item -LiteralPath $SourcePath
    if (-not $directory.PSIsContainer) {
        $directory = Get-Item -LiteralPath (Split-Path -Path $directory.FullName -Parent)
    }

    if ($directory.Name -notin @("Light", "Lights", "Good", "Trash")) {
        return $directory.Name
    }

    return ""
}

function Get-AsiToPixFitsFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath
    )

    return @(Get-ChildItem -LiteralPath $SourcePath -File -Recurse -ErrorAction Stop |
        Where-Object { $_.Name -match '\.fits?(\.gz)?$' } |
        Sort-Object FullName)
}

function Resolve-AsiToPixSeasonName {
    param(
        [string]$ProvidedSeasonName = "",

        [string[]]$ExistingSeasonNames = @()
    )

    if (-not [string]::IsNullOrWhiteSpace($ProvidedSeasonName)) {
        return $ProvidedSeasonName
    }

    if ($ExistingSeasonNames.Count -eq 1) {
        $candidate = $ExistingSeasonNames[0]
        if (Read-AsiToPixConfirmation -Prompt "Use existing season/group '$candidate'?") {
            return $candidate
        }
    } elseif ($ExistingSeasonNames.Count -gt 1) {
        Write-Host "`nExisting destination seasons/groups:" -ForegroundColor Cyan
        for ($i = 0; $i -lt $ExistingSeasonNames.Count; $i++) {
            Write-Host " [$($i + 1)] $($ExistingSeasonNames[$i])" -ForegroundColor White
        }
        Write-Host " [0] Use a new name" -ForegroundColor White

        do {
            $answer = (Read-Host "Select season/group index, or type a new name").Trim()
            $selectedIndex = -1
            if ([int]::TryParse($answer, [ref]$selectedIndex)) {
                if ($selectedIndex -eq 0) { break }
                if ($selectedIndex -ge 1 -and $selectedIndex -le $ExistingSeasonNames.Count) {
                    return $ExistingSeasonNames[$selectedIndex - 1]
                }
            } elseif (-not [string]::IsNullOrWhiteSpace($answer)) {
                return $answer
            }

            Write-Host "[!] Invalid selection." -ForegroundColor Red
        } while ($true)
    }

    $defaultSeason = (Get-Date).Year.ToString([System.Globalization.CultureInfo]::InvariantCulture)
    return (Read-AsiToPixRequiredValue -Prompt "Enter destination season/group name" -DefaultValue $defaultSeason)
}

function Resolve-AsiToPixSetupName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$DetectedTelescopeName,

        [AllowEmptyString()]
        [string]$DetectedCameraName,

        [string]$ProvidedTelescopeName = "",

        [string]$ProvidedCameraName = "",

        [string[]]$ExistingSetupNames = @()
    )

    $cameraName = if ([string]::IsNullOrWhiteSpace($ProvidedCameraName)) { $DetectedCameraName } else { $ProvidedCameraName }
    if ([string]::IsNullOrWhiteSpace($cameraName)) {
        $cameraName = Read-AsiToPixRequiredValue -Prompt "Enter destination camera name (for example ASI2600MM)"
    }

    if (-not [string]::IsNullOrWhiteSpace($ProvidedTelescopeName)) {
        return "$ProvidedTelescopeName @ $cameraName"
    }

    $selection = Read-AsiToPixNameSelection -Kind "telescope/setup" -DetectedName $DetectedTelescopeName -Candidates $ExistingSetupNames

    if (-not $selection.IsNew) {
        return $selection.Name
    }

    $selectedSetup = Get-AsiToPixSetupInfo -SetupName $selection.Name
    if (-not [string]::IsNullOrWhiteSpace($selectedSetup.Camera)) {
        return $selection.Name
    }

    return "$($selection.Name) @ $cameraName"
}

function New-AsiToPixDirectory {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [System.Management.Automation.PSCmdlet]$Cmdlet
    )

    if (Test-Path -LiteralPath $Path -PathType Container) {
        return
    }

    if ($Cmdlet.ShouldProcess($Path, "Create directory")) {
        New-Item -ItemType Directory -Path $Path -Force -ErrorAction Stop | Out-Null
    }
}

function Get-AsiToPixFileNameSet {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath
    )

    $set = [System.Collections.Generic.HashSet[string]]::new([System.StringComparer]::OrdinalIgnoreCase)
    if (Test-Path -LiteralPath $RootPath -PathType Container) {
        foreach ($file in Get-ChildItem -LiteralPath $RootPath -File -Recurse -ErrorAction Stop) {
            [void]$set.Add($file.Name)
        }
    }

    return ,$set
}

function Import-AsiToPixSession {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,

        [Parameter(Mandatory = $true)]
        [string]$AstroPhotoRoot,

        [string]$ObjectName = "",

        [string]$SeasonName = "",

        [string]$TelescopeName = "",

        [string]$CameraName = ""
    )

    $resolvedSourcePath = (Resolve-Path -LiteralPath $SourcePath).ProviderPath
    $resolvedAstroPhotoRoot = (Resolve-Path -LiteralPath $AstroPhotoRoot).ProviderPath

    $files = @(Get-AsiToPixFitsFile -SourcePath $resolvedSourcePath)
    if ($files.Count -eq 0) {
        throw "No FITS files found under source path: $resolvedSourcePath"
    }

    $parsedFiles = foreach ($file in $files) {
        $info = Get-AsiToPixLightFileInfo -FileName $file.Name
        if ($null -eq $info.CapturedAt) {
            Write-Host "[!] Skipping file without ASIAir timestamp: $($file.FullName)" -ForegroundColor Yellow
            continue
        }

        [PSCustomObject]@{
            File        = $file
            ObjectName  = $info.ObjectName
            CameraName  = $info.CameraName
            FilterName  = $info.FilterName
            CapturedAt  = $info.CapturedAt
            NightDate   = Get-AsiToPixNightDate -CapturedAt $info.CapturedAt
            Telescope   = $info.TelescopeName
        }
    }

    $parsedFiles = @($parsedFiles)
    if ($parsedFiles.Count -eq 0) {
        throw "No importable ASIAir light files found under source path: $resolvedSourcePath"
    }

    $sample = $parsedFiles[0]
    $detectedSourceObject = Get-AsiToPixDetectedObject -SourcePath $resolvedSourcePath
    $detectedObject = if ([string]::IsNullOrWhiteSpace($ObjectName)) { $detectedSourceObject } else { $ObjectName }
    $detectedCamera = if ([string]::IsNullOrWhiteSpace($CameraName)) { $sample.CameraName } else { $CameraName }
    $detectedTelescope = Get-AsiToPixDetectedTelescope -SourcePath $resolvedSourcePath -FallbackTelescope $sample.Telescope

    if ([string]::IsNullOrWhiteSpace($detectedObject)) {
        $detectedObject = Read-AsiToPixRequiredValue -Prompt "Enter detected/source object name"
    }

    if ([string]::IsNullOrWhiteSpace($detectedTelescope)) {
        $detectedTelescope = Read-AsiToPixRequiredValue -Prompt "Enter detected/source telescope name"
    }

    $asiairRoot = Join-Path -Path $resolvedAstroPhotoRoot -ChildPath "ASIAir"
    $existingObjects = @()
    if (Test-Path -LiteralPath $asiairRoot -PathType Container) {
        $existingObjects = @(Get-ChildItem -LiteralPath $asiairRoot -Directory | Select-Object -ExpandProperty Name)
    }

    $objectSelection = if ([string]::IsNullOrWhiteSpace($ObjectName)) {
        Read-AsiToPixNameSelection -Kind "object" -DetectedName $detectedObject -Candidates $existingObjects
    } else {
        [PSCustomObject]@{
            Name  = $ObjectName
            IsNew = ($existingObjects -notcontains $ObjectName)
        }
    }

    $objectPath = Join-Path -Path $asiairRoot -ChildPath $objectSelection.Name
    $existingSeasons = @()
    if (Test-Path -LiteralPath $objectPath -PathType Container) {
        $existingSeasons = @(Get-ChildItem -LiteralPath $objectPath -Directory | Select-Object -ExpandProperty Name)
    }

    $resolvedSeasonName = Resolve-AsiToPixSeasonName -ProvidedSeasonName $SeasonName -ExistingSeasonNames $existingSeasons
    $seasonPath = Join-Path -Path $objectPath -ChildPath $resolvedSeasonName

    $existingSetups = @()
    if (Test-Path -LiteralPath $seasonPath -PathType Container) {
        $existingSetups = @(Get-ChildItem -LiteralPath $seasonPath -Directory | Select-Object -ExpandProperty Name)
    }

    $setupName = Resolve-AsiToPixSetupName `
        -DetectedTelescopeName $detectedTelescope `
        -DetectedCameraName $detectedCamera `
        -ProvidedTelescopeName $TelescopeName `
        -ProvidedCameraName $CameraName `
        -ExistingSetupNames $existingSetups

    $setupPath = Join-Path -Path $seasonPath -ChildPath $setupName
    $goodRoot = Join-Path -Path $setupPath -ChildPath "Good"
    $trashRoot = Join-Path -Path $setupPath -ChildPath "Trash"

    $summaryByFilterAndNight = $parsedFiles |
        Group-Object -Property FilterName, NightDate |
        Sort-Object Name

    Write-Host "`nImport target:" -ForegroundColor Cyan
    Write-Host "  Object:    $($objectSelection.Name)" -ForegroundColor White
    Write-Host "  Season:    $resolvedSeasonName" -ForegroundColor White
    Write-Host "  Setup:     $setupName" -ForegroundColor White
    Write-Host "  Root:      $setupPath" -ForegroundColor White
    Write-Host "`nDetected groups:" -ForegroundColor Cyan
    foreach ($group in $summaryByFilterAndNight) {
        $first = $group.Group[0]
        Write-Host "  $($first.FilterName) / $($first.NightDate): $($group.Count) file(s)" -ForegroundColor White
    }

    if (-not (Read-AsiToPixConfirmation -Prompt "Merge/copy this session into the target tree?")) {
        Write-Host "[INFO] Import cancelled." -ForegroundColor Yellow
        return
    }

    New-AsiToPixDirectory -Path $asiairRoot -Cmdlet $PSCmdlet
    New-AsiToPixDirectory -Path $objectPath -Cmdlet $PSCmdlet
    New-AsiToPixDirectory -Path $seasonPath -Cmdlet $PSCmdlet
    New-AsiToPixDirectory -Path $setupPath -Cmdlet $PSCmdlet
    New-AsiToPixDirectory -Path $goodRoot -Cmdlet $PSCmdlet
    New-AsiToPixDirectory -Path $trashRoot -Cmdlet $PSCmdlet

    $goodNames = Get-AsiToPixFileNameSet -RootPath $goodRoot
    $trashNames = Get-AsiToPixFileNameSet -RootPath $trashRoot

    $copied = 0
    $skippedExisting = 0
    $skippedTrash = 0

    foreach ($entry in $parsedFiles) {
        $fileName = $entry.File.Name

        if ($trashNames.Contains($fileName)) {
            $skippedTrash++
            Write-Host "  [trash] $fileName" -ForegroundColor DarkYellow
            continue
        }

        if ($goodNames.Contains($fileName)) {
            $skippedExisting++
            Write-Host "  [exists] $fileName" -ForegroundColor DarkGray
            continue
        }

        $filterGoodPath = Join-Path -Path $goodRoot -ChildPath $entry.FilterName
        $filterTrashPath = Join-Path -Path $trashRoot -ChildPath $entry.FilterName
        $nightGoodPath = Join-Path -Path $filterGoodPath -ChildPath $entry.NightDate
        $nightTrashPath = Join-Path -Path $filterTrashPath -ChildPath $entry.NightDate
        $destinationFile = Join-Path -Path $nightGoodPath -ChildPath $fileName

        New-AsiToPixDirectory -Path $filterGoodPath -Cmdlet $PSCmdlet
        New-AsiToPixDirectory -Path $filterTrashPath -Cmdlet $PSCmdlet
        New-AsiToPixDirectory -Path $nightGoodPath -Cmdlet $PSCmdlet
        New-AsiToPixDirectory -Path $nightTrashPath -Cmdlet $PSCmdlet

        if ($PSCmdlet.ShouldProcess($destinationFile, "Copy FITS file from '$($entry.File.FullName)'")) {
            Copy-Item -LiteralPath $entry.File.FullName -Destination $destinationFile -ErrorAction Stop
            [void]$goodNames.Add($fileName)
            $copied++
            Write-Host "  [+] $fileName -> $($entry.FilterName)\$($entry.NightDate)" -ForegroundColor Green
        }
    }

    Write-Host "`n[DONE] Import finished." -ForegroundColor Cyan
    Write-Host "  Copied:           $copied" -ForegroundColor White
    Write-Host "  Already in Good:  $skippedExisting" -ForegroundColor White
    Write-Host "  Preserved Trash:  $skippedTrash" -ForegroundColor White
}

Export-ModuleMember -Function `
    Get-AsiToPixLightFileInfo, `
    Get-AsiToPixNameMatch, `
    Get-AsiToPixNightDate, `
    Get-AsiToPixSetupInfo, `
    Import-AsiToPixSession
