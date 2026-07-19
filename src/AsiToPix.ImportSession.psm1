Set-StrictMode -Version Latest

$imageFilesModule = Join-Path -Path $PSScriptRoot -ChildPath "AsiToPix.ImageFiles.psm1"
Import-Module $imageFilesModule -Force

$frameFoldersModule = Join-Path -Path $PSScriptRoot -ChildPath "AsiToPix.FrameFolders.psm1"
Import-Module $frameFoldersModule -Force

$namesModule = Join-Path -Path $PSScriptRoot -ChildPath "AsiToPix.Names.psm1"
Import-Module $namesModule -Force

function ConvertTo-AsiToPixPathSegment {
    param(
        [AllowEmptyString()]
        [string]$Value = "",

        [string]$ValueName = "path segment",

        [switch]$Quiet
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return ""
    }

    $originalValue = $Value
    $resolvedValue = $Value.Normalize([System.Text.NormalizationForm]::FormC)
    $resolvedValue = [regex]::Replace($resolvedValue, '[\p{Cc}\p{Cf}]', ' ')

    foreach ($invalidChar in [System.IO.Path]::GetInvalidFileNameChars()) {
        $resolvedValue = $resolvedValue.Replace($invalidChar, [char]' ')
    }

    $resolvedValue = [regex]::Replace($resolvedValue, '\s+', ' ').Trim()
    $resolvedValue = $resolvedValue.TrimEnd([char[]]@('.', ' '))

    if ([string]::IsNullOrWhiteSpace($resolvedValue)) {
        throw "The $ValueName '$originalValue' does not contain a valid Windows path segment after sanitizing."
    }

    if (-not $Quiet -and $resolvedValue -ne $originalValue) {
        Write-Host "[INFO] Sanitized ${ValueName}: '$originalValue' -> '$resolvedValue'" -ForegroundColor Yellow
    }

    return $resolvedValue
}

function Read-AsiToPixRequiredValue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Prompt,

        [string]$DefaultValue = "",

        [string]$ValueName = "value"
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
            return (ConvertTo-AsiToPixPathSegment -Value $value -ValueName $ValueName)
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

function Read-AsiToPixImportMode {
    param(
        [ValidateSet("", "Copy", "Symlink")]
        [string]$ImportMode = ""
    )

    if (-not [string]::IsNullOrWhiteSpace($ImportMode)) {
        return $ImportMode
    }

    Write-Host "`nImport mode:" -ForegroundColor Cyan
    Write-Host " [1] Copy (default)" -ForegroundColor White
    Write-Host " [2] Symlink" -ForegroundColor White

    do {
        $answer = (Read-Host "Select import mode").Trim()
        if ([string]::IsNullOrWhiteSpace($answer)) {
            return "Copy"
        }

        switch -Regex ($answer) {
            '^(1|copy|c)$' { return "Copy" }
            '^(2|symlink|link|s)$' { return "Symlink" }
        }

        Write-Host "[!] Invalid import mode. Enter 1/Copy or 2/Symlink." -ForegroundColor Red
    } while ($true)
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

        $newName = Read-AsiToPixRequiredValue -Prompt "Enter destination $Kind name" -DefaultValue $DetectedName -ValueName "$Kind name"
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
                    $newName = Read-AsiToPixRequiredValue -Prompt "Enter destination $Kind name" -DefaultValue $DetectedName -ValueName "$Kind name"
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
                $answer = ConvertTo-AsiToPixPathSegment -Value $answer -ValueName "$Kind name"
                return [PSCustomObject]@{
                    Name  = $answer
                    IsNew = ($Candidates -notcontains $answer)
                }
            }

            Write-Host "[!] Invalid selection." -ForegroundColor Red
        } while ($true)
    }

    Write-Host "[INFO] No matching existing $Kind found for '$DetectedName'." -ForegroundColor Yellow
    $value = Read-AsiToPixRequiredValue -Prompt "Enter destination $Kind name" -DefaultValue $DetectedName -ValueName "$Kind name"

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

    $nameWithoutExtension = Get-AsiToPixImageFileStem -FileName $FileName

    $objectName = $null
    $exposureSeconds = $null
    $exposureUnit = $null
    if ($nameWithoutExtension -match '^Light_(?<object>.+?)_(?<exp>\d+(?:\.\d+)?)(?<unit>m?s)(?:_|$)') {
        $objectName = ($Matches["object"] -replace '_', ' ').Trim()
        $exposureSeconds = $Matches["exp"]
        $exposureUnit = $Matches["unit"]
    } elseif ($nameWithoutExtension -match '^(?:Dark|Bias|Flat)_(?<exp>\d+(?:\.\d+)?)(?<unit>m?s)(?:_|$)') {
        $exposureSeconds = $Matches["exp"]
        $exposureUnit = $Matches["unit"]
    }

    if ($null -ne $exposureSeconds -and $exposureUnit -eq "ms") {
        $exposureMilliseconds = [decimal]::Parse(
            $exposureSeconds,
            [System.Globalization.NumberStyles]::Number,
            [System.Globalization.CultureInfo]::InvariantCulture
        )
        $exposureSeconds = ($exposureMilliseconds / 1000).ToString(
            "0.############################",
            [System.Globalization.CultureInfo]::InvariantCulture
        )
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

function Resolve-AsiToPixLightFileInfo {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileName,

        [string]$FilterName = "",

        [string]$ExposureSeconds = "",

        [switch]$PromptForMissingData
    )

    $info = Get-AsiToPixLightFileInfo -FileName $FileName
    if ($null -eq $info.ExposureSeconds) {
        $resolvedExposureSeconds = $ExposureSeconds.Trim().Replace(',', '.')
        if ([string]::IsNullOrWhiteSpace($resolvedExposureSeconds) -and $PromptForMissingData) {
            $resolvedExposureSeconds = Read-AsiToPixRequiredValue `
                -Prompt "Enter the light exposure in seconds for file '$FileName'" `
                -ValueName "exposure"
            $resolvedExposureSeconds = $resolvedExposureSeconds.Replace(',', '.')
        }
        if (-not [string]::IsNullOrWhiteSpace($resolvedExposureSeconds)) {
            $exposureValue = [decimal]0
            if (-not [decimal]::TryParse(
                $resolvedExposureSeconds,
                [System.Globalization.NumberStyles]::Float,
                [System.Globalization.CultureInfo]::InvariantCulture,
                [ref]$exposureValue
            ) -or $exposureValue -le 0) {
                throw "Invalid light exposure '$resolvedExposureSeconds' for file: $FileName"
            }
            $info.ExposureSeconds = $exposureValue.ToString(
                "0.############################",
                [System.Globalization.CultureInfo]::InvariantCulture
            )
        }
    }

    if ($info.CameraName -match 'MM$' -and $info.FilterName -eq "None") {
        $resolvedFilterName = $FilterName.Trim()
        if ([string]::IsNullOrWhiteSpace($resolvedFilterName) -and $PromptForMissingData) {
            $resolvedFilterName = Read-AsiToPixRequiredValue `
                -Prompt "Enter the light filter for file '$FileName' (for example L, H, O, or S)" `
                -ValueName "filter name"
        }
        if (-not [string]::IsNullOrWhiteSpace($resolvedFilterName)) {
            $info.FilterName = $resolvedFilterName
        }
    }
    $info.FilterName = ConvertTo-AsiToPixDestinationFilterName -FilterName $info.FilterName -CameraName $info.CameraName

    return $info
}

function Get-AsiToPixCameraBaseName {
    param(
        [AllowEmptyString()]
        [string]$CameraName = ""
    )

    $trimmedName = $CameraName.Trim()
    if ([string]::IsNullOrWhiteSpace($trimmedName)) {
        return ""
    }

    if ($trimmedName -match '^(?:ASI)?(?<number>\d{3,4})(?<suffix>MM|MC|M)?$') {
        return "ASI$($Matches["number"])"
    }

    return $trimmedName
}

function Get-AsiToPixCameraType {
    param(
        [AllowEmptyString()]
        [string]$CameraName = ""
    )

    $trimmedName = $CameraName.Trim()
    if ($trimmedName -match 'MM$') {
        return "Mono"
    }

    if ($trimmedName -match 'MC$') {
        return "OSC"
    }

    return "Unknown"
}

function ConvertTo-AsiToPixSetupCameraName {
    param(
        [AllowEmptyString()]
        [string]$SetupName = ""
    )

    if ([string]::IsNullOrWhiteSpace($SetupName)) {
        return $SetupName
    }

    return [regex]::Replace(
        $SetupName,
        '\b(?:ASI)?\d{3,4}(?:MM|MC|M)?\b',
        {
            param($match)
            Get-AsiToPixCameraBaseName -CameraName $match.Value
        }
    )
}

function ConvertTo-AsiToPixDestinationFilterName {
    param(
        [AllowEmptyString()]
        [string]$FilterName = "",

        [AllowEmptyString()]
        [string]$CameraName = ""
    )

    $resolvedFilterName = $FilterName.Trim()
    if ([string]::IsNullOrWhiteSpace($resolvedFilterName)) {
        $resolvedFilterName = "None"
    }

    $cameraType = Get-AsiToPixCameraType -CameraName $CameraName
    if ($cameraType -eq "OSC") {
        switch -Regex ($resolvedFilterName) {
            '^(None|RGB)$' { return "RGB" }
            '^(HO|UHC)$' { return "HO" }
            '^SO$' { return "SO" }
            default { return $resolvedFilterName }
        }
    }

    switch -Regex ($resolvedFilterName) {
        '^(Ha)$' { return "H" }
        '^(SII)$' { return "S" }
        '^(OIII|OII)$' { return "O" }
        default { return $resolvedFilterName }
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

function ConvertTo-AsiToPixExposureFolderSuffix {
    param(
        [AllowEmptyString()]
        [string]$ExposureSeconds = ""
    )

    if ([string]::IsNullOrWhiteSpace($ExposureSeconds)) {
        return "unknown"
    }

    try {
        $exposureValue = [decimal]::Parse(
            $ExposureSeconds,
            [System.Globalization.NumberStyles]::Number,
            [System.Globalization.CultureInfo]::InvariantCulture
        )

        return "$($exposureValue.ToString('0.########', [System.Globalization.CultureInfo]::InvariantCulture))s"
    } catch {
        return (ConvertTo-AsiToPixPathSegment -Value $ExposureSeconds -ValueName "exposure suffix" -Quiet)
    }
}

function Get-AsiToPixDestinationNightFolder {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Entry
    )

    $property = $Entry.PSObject.Properties["DestinationNightFolder"]
    if ($property -and -not [string]::IsNullOrWhiteSpace($property.Value)) {
        return $property.Value
    }

    return $Entry.NightDate
}

function Test-AsiToPixRawFileName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileName
    )

    return (Test-AsiToPixRawImageFileName -FileName $FileName)
}

function Test-AsiToPixSupportedLightFileName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileName
    )

    return (Test-AsiToPixSupportedImageFileName -FileName $FileName)
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
        if ((Test-AsiToPixFrameFolderName -Name $current.Name -Kind Light) -and $null -ne $current.Parent) {
            return $current.Parent.Name
        }

        if ($null -ne $current.Parent -and (Test-AsiToPixFrameFolderName -Name $current.Parent.Name -Kind Light)) {
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

    if (-not (Test-AsiToPixFrameFolderName -Name $directory.Name -Kind Light) -and
        $directory.Name -notin @("Good", "Trash")) {
        return $directory.Name
    }

    return ""
}

function Get-AsiToPixSourceLightFile {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath
    )

    return @(Get-ChildItem -LiteralPath $SourcePath -File -Recurse -ErrorAction Stop |
        Where-Object { Test-AsiToPixSupportedLightFileName -FileName $_.Name } |
        Sort-Object FullName)
}

function Resolve-AsiToPixSeasonName {
    param(
        [string]$ProvidedSeasonName = "",

        [string[]]$ExistingSeasonNames = @()
    )

    if (-not [string]::IsNullOrWhiteSpace($ProvidedSeasonName)) {
        return (ConvertTo-AsiToPixPathSegment -Value $ProvidedSeasonName -ValueName "season/group name")
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
                $answer = ConvertTo-AsiToPixPathSegment -Value $answer -ValueName "season/group name"
                return $answer
            }

            Write-Host "[!] Invalid selection." -ForegroundColor Red
        } while ($true)
    }

    $defaultSeason = (Get-Date).Year.ToString([System.Globalization.CultureInfo]::InvariantCulture)
    return (Read-AsiToPixRequiredValue -Prompt "Enter destination season/group name" -DefaultValue $defaultSeason -ValueName "season/group name")
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
    $cameraName = ConvertTo-AsiToPixPathSegment -Value $cameraName -ValueName "camera name"
    $cameraName = Get-AsiToPixCameraBaseName -CameraName $cameraName
    if ([string]::IsNullOrWhiteSpace($cameraName)) {
        $cameraName = Read-AsiToPixRequiredValue -Prompt "Enter destination camera name without MM/MC suffix (for example ASI2600)" -ValueName "camera name"
        $cameraName = Get-AsiToPixCameraBaseName -CameraName $cameraName
    }

    if (-not [string]::IsNullOrWhiteSpace($ProvidedTelescopeName)) {
        $telescopeName = ConvertTo-AsiToPixPathSegment -Value $ProvidedTelescopeName -ValueName "telescope/setup name"
        return (ConvertTo-AsiToPixPathSegment -Value "$telescopeName @ $cameraName" -ValueName "setup name")
    }

    $selection = Read-AsiToPixNameSelection -Kind "telescope/setup" -DetectedName $DetectedTelescopeName -Candidates $ExistingSetupNames

    if (-not $selection.IsNew) {
        return (ConvertTo-AsiToPixSetupCameraName -SetupName $selection.Name)
    }

    $selectedSetup = Get-AsiToPixSetupInfo -SetupName $selection.Name
    if (-not [string]::IsNullOrWhiteSpace($selectedSetup.Camera)) {
        return (ConvertTo-AsiToPixSetupCameraName -SetupName $selection.Name)
    }

    return (ConvertTo-AsiToPixPathSegment -Value "$($selection.Name) @ $cameraName" -ValueName "setup name")
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

function Find-AsiToPixImportSession {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ImportRoot
    )

    $resolvedImportRoot = (Resolve-Path -LiteralPath $ImportRoot).ProviderPath
    $sessions = @()

    foreach ($setupFolder in Get-ChildItem -LiteralPath $resolvedImportRoot -Directory -ErrorAction Stop) {
        $lightFolders = @(Get-AsiToPixChildFrameFolder -Path $setupFolder.FullName -Kind Light)

        foreach ($lightFolder in $lightFolders) {
            $objectFolders = @(Get-ChildItem -LiteralPath $lightFolder.FullName -Directory -ErrorAction Stop)

            foreach ($objectFolder in $objectFolders) {
                $files = @(Get-AsiToPixSourceLightFile -SourcePath $objectFolder.FullName)
                if ($files.Count -eq 0) {
                    continue
                }

                $sessions += [PSCustomObject]@{
                    SourcePath        = $objectFolder.FullName
                    SetupSourcePath   = $setupFolder.FullName
                    DetectedSetupName = $setupFolder.Name
                    DetectedObject    = $objectFolder.Name
                    FileCount         = $files.Count
                }
            }
        }
    }

    return @($sessions | Sort-Object DetectedSetupName, DetectedObject, SourcePath)
}

function Get-AsiToPixDefaultImportRoot {
    param(
        [string]$AstroPhotoRoot = ""
    )

    $roots = @()
    if (-not [string]::IsNullOrWhiteSpace($AstroPhotoRoot)) {
        $resolvedAstroPhotoRoot = (Resolve-Path -LiteralPath $AstroPhotoRoot).ProviderPath
        $importPath = Join-Path -Path $resolvedAstroPhotoRoot -ChildPath "Import"
        if (Test-Path -LiteralPath $importPath -PathType Container) {
            $roots += (Resolve-Path -LiteralPath $importPath).ProviderPath
        }

        return @($roots | Sort-Object -Unique)
    }

    foreach ($drive in Get-PSDrive -PSProvider FileSystem) {
        $astroPhotoPath = Join-Path -Path $drive.Root -ChildPath "AstroPhoto"
        $importPath = Join-Path -Path $astroPhotoPath -ChildPath "Import"
        if (Test-Path -LiteralPath $importPath -PathType Container -ErrorAction SilentlyContinue) {
            $roots += (Resolve-Path -LiteralPath $importPath -ErrorAction Stop).ProviderPath
        }
    }

    return @($roots | Sort-Object -Unique)
}

function Resolve-AsiToPixImportSourcePath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,

        [string]$AstroPhotoRoot = ""
    )

    $resolvedInput = $SourcePath.Trim().Trim('"')
    if ($resolvedInput -match '(?i)\.(fits?|fits?\.gz|arw|cr2|cr3|nef|nrw|raf|orf|rw2|dng|pef|srw|3fr|erf|kdc|mos|mrw|raw)$') {
        $resolvedInput = Split-Path -Path $resolvedInput -Parent
    }

    if (Test-Path -LiteralPath $resolvedInput -PathType Container) {
        return [PSCustomObject]@{
            SourcePath     = (Resolve-Path -LiteralPath $resolvedInput).ProviderPath
            AstroPhotoRoot = $AstroPhotoRoot
        }
    }

    if ([System.IO.Path]::IsPathRooted($resolvedInput) -or $resolvedInput -match '[\\/]') {
        return [PSCustomObject]@{
            SourcePath     = $resolvedInput
            AstroPhotoRoot = $AstroPhotoRoot
        }
    }

    $importRoots = @(Get-AsiToPixDefaultImportRoot -AstroPhotoRoot $AstroPhotoRoot)
    if ($importRoots.Count -eq 0) {
        throw "No Import folder found by pattern *:\AstroPhoto\Import. Enter a full source path instead of object name '$resolvedInput'."
    }

    $sessions = foreach ($importRoot in $importRoots) {
        Find-AsiToPixImportSession -ImportRoot $importRoot
    }
    $sessions = @($sessions)
    if ($sessions.Count -eq 0) {
        throw "No import sessions with supported light files found under Import folder(s): $($importRoots -join ', ')"
    }

    $objectNames = @($sessions | Select-Object -ExpandProperty DetectedObject -Unique)
    $objectMatches = @(Get-AsiToPixNameMatch -DetectedName $resolvedInput -Candidates $objectNames)
    if ($objectMatches.Count -eq 0) {
        throw "No import object folder matching '$resolvedInput' found under Import folder(s): $($importRoots -join ', ')"
    }

    $matchedNames = @($objectMatches | Select-Object -ExpandProperty Name)
    $candidateSessions = @($sessions |
        Where-Object { $matchedNames -contains $_.DetectedObject } |
        Sort-Object @{ Expression = { [array]::IndexOf($matchedNames, $_.DetectedObject) } }, DetectedSetupName, SourcePath)

    if ($candidateSessions.Count -eq 1) {
        $selected = $candidateSessions[0]
        Write-Host "[INFO] Import object '$resolvedInput' resolved to: $($selected.SourcePath)" -ForegroundColor Cyan
    } else {
        Write-Host "`nMatching import sessions for '${resolvedInput}':" -ForegroundColor Cyan
        for ($i = 0; $i -lt $candidateSessions.Count; $i++) {
            $candidate = $candidateSessions[$i]
            Write-Host (" [{0}] {1} | {2} | {3} file(s)" -f ($i + 1), $candidate.DetectedSetupName, $candidate.DetectedObject, $candidate.FileCount) -ForegroundColor White
            Write-Host "     $($candidate.SourcePath)" -ForegroundColor DarkGray
        }

        do {
            $answer = (Read-Host "Select import session index, press Enter for 1, or type a full source path").Trim()
            if ([string]::IsNullOrWhiteSpace($answer)) {
                $selected = $candidateSessions[0]
                break
            }

            if ($answer -match '^\d+$') {
                $index = [int]$answer - 1
                if ($index -ge 0 -and $index -lt $candidateSessions.Count) {
                    $selected = $candidateSessions[$index]
                    break
                }
            } elseif (Test-Path -LiteralPath $answer -PathType Container) {
                return [PSCustomObject]@{
                    SourcePath     = (Resolve-Path -LiteralPath $answer).ProviderPath
                    AstroPhotoRoot = $AstroPhotoRoot
                }
            }

            Write-Host "[!] Invalid import session selection." -ForegroundColor Red
        } while ($true)
    }

    $selectedImportRoot = $importRoots |
        Where-Object { $selected.SourcePath.StartsWith($_, [System.StringComparison]::OrdinalIgnoreCase) } |
        Sort-Object Length -Descending |
        Select-Object -First 1
    $selectedAstroPhotoRoot = if ($selectedImportRoot) {
        Split-Path -Path $selectedImportRoot -Parent
    } else {
        $AstroPhotoRoot
    }

    return [PSCustomObject]@{
        SourcePath     = $selected.SourcePath
        AstroPhotoRoot = $selectedAstroPhotoRoot
    }
}

function Get-AsiToPixImportPlan {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,

        [Parameter(Mandatory = $true)]
        [string]$AstroPhotoRoot,

        [string]$ObjectName = "",

        [string]$SeasonName = "",

        [string]$SetupName = "",

        [string]$TelescopeName = "",

        [string]$CameraName = "",

        [ValidateSet("", "Copy", "Symlink")]
        [string]$ImportMode = ""
    )

    $resolvedSourcePath = (Resolve-Path -LiteralPath $SourcePath).ProviderPath
    $resolvedAstroPhotoRoot = (Resolve-Path -LiteralPath $AstroPhotoRoot).ProviderPath
    $ObjectName = ConvertTo-AsiToPixPathSegment -Value $ObjectName -ValueName "object name"
    $SeasonName = ConvertTo-AsiToPixPathSegment -Value $SeasonName -ValueName "season/group name"
    $SetupName = ConvertTo-AsiToPixPathSegment -Value $SetupName -ValueName "setup name"
    $TelescopeName = ConvertTo-AsiToPixPathSegment -Value $TelescopeName -ValueName "telescope/setup name"
    $CameraName = ConvertTo-AsiToPixPathSegment -Value $CameraName -ValueName "camera name"

    $files = @(Get-AsiToPixSourceLightFile -SourcePath $resolvedSourcePath)
    if ($files.Count -eq 0) {
        throw "No supported PixInsight image files found under source path: $resolvedSourcePath"
    }

    $detectedSourceObject = Get-AsiToPixDetectedObject -SourcePath $resolvedSourcePath
    $resolvedMissingFilter = ""
    $resolvedMissingExposure = ""
    $parsedFiles = foreach ($file in $files) {
        $originalInfo = Get-AsiToPixLightFileInfo -FileName $file.Name
        $info = Resolve-AsiToPixLightFileInfo `
            -FileName $file.Name `
            -FilterName $resolvedMissingFilter `
            -ExposureSeconds $resolvedMissingExposure `
            -PromptForMissingData
        if ($null -eq $originalInfo.ExposureSeconds -and $null -ne $info.ExposureSeconds -and
            [string]::IsNullOrWhiteSpace($resolvedMissingExposure)) {
            $resolvedMissingExposure = $info.ExposureSeconds
        }
        if ($info.CameraName -match 'MM$' -and $info.FilterName -ne "None" -and
            [string]::IsNullOrWhiteSpace($resolvedMissingFilter) -and
            $originalInfo.FilterName -eq "None") {
            $resolvedMissingFilter = $info.FilterName
        }
        $capturedAt = $info.CapturedAt
        if ($null -eq $capturedAt) {
            $capturedAt = $file.LastWriteTime
        }

        if ($null -eq $capturedAt) {
            Write-Host "[!] Skipping file without importable timestamp: $($file.FullName)" -ForegroundColor Yellow
            continue
        }

        [PSCustomObject]@{
            File        = $file
            ObjectName  = $info.ObjectName
            CameraName  = $info.CameraName
            FilterName  = $info.FilterName
            ExposureSeconds = $info.ExposureSeconds
            CapturedAt  = $capturedAt
            NightDate   = Get-AsiToPixNightDate -CapturedAt $capturedAt
            Telescope   = $info.TelescopeName
        }
    }

    $parsedFiles = @($parsedFiles)
    if ($parsedFiles.Count -eq 0) {
        throw "No importable light image files found under source path: $resolvedSourcePath"
    }

    $parsedFilesByFilterAndNight = $parsedFiles | Group-Object -Property FilterName, NightDate
    foreach ($group in $parsedFilesByFilterAndNight) {
        $exposureSuffixes = @($group.Group |
            ForEach-Object { ConvertTo-AsiToPixExposureFolderSuffix -ExposureSeconds $_.ExposureSeconds } |
            Sort-Object -Unique)
        $useExposureSuffix = ($exposureSuffixes.Count -gt 1)

        foreach ($entry in $group.Group) {
            $destinationNightFolder = $entry.NightDate
            if ($useExposureSuffix) {
                $destinationNightFolder = "$($entry.NightDate)-$(ConvertTo-AsiToPixExposureFolderSuffix -ExposureSeconds $entry.ExposureSeconds)"
            }

            $entry | Add-Member -NotePropertyName DestinationNightFolder -NotePropertyValue $destinationNightFolder -Force
        }
    }

    $sample = $parsedFiles[0]
    $detectedSourceObject = ConvertTo-AsiToPixPathSegment -Value $detectedSourceObject -ValueName "detected object name"
    $detectedObject = if ([string]::IsNullOrWhiteSpace($ObjectName)) { $detectedSourceObject } else { $ObjectName }
    $detectedCamera = if ([string]::IsNullOrWhiteSpace($CameraName)) { $sample.CameraName } else { $CameraName }
    $detectedTelescope = Get-AsiToPixDetectedTelescope -SourcePath $resolvedSourcePath -FallbackTelescope $sample.Telescope
    $detectedTelescope = ConvertTo-AsiToPixPathSegment -Value $detectedTelescope -ValueName "detected telescope/setup name"

    if ([string]::IsNullOrWhiteSpace($detectedObject)) {
        $detectedObject = Read-AsiToPixRequiredValue -Prompt "Enter detected/source object name" -ValueName "detected object name"
    }

    if ([string]::IsNullOrWhiteSpace($detectedTelescope)) {
        $detectedTelescope = Read-AsiToPixRequiredValue -Prompt "Enter detected/source telescope name" -ValueName "detected telescope/setup name"
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

    $resolvedSetupName = if ([string]::IsNullOrWhiteSpace($SetupName)) {
        Resolve-AsiToPixSetupName `
            -DetectedTelescopeName $detectedTelescope `
            -DetectedCameraName $detectedCamera `
            -ProvidedTelescopeName $TelescopeName `
            -ProvidedCameraName $CameraName `
            -ExistingSetupNames $existingSetups
    } else {
        ConvertTo-AsiToPixSetupCameraName -SetupName $SetupName
    }

    $setupPath = Join-Path -Path $seasonPath -ChildPath $resolvedSetupName
    $goodRoot = Join-Path -Path $setupPath -ChildPath "Good"
    $trashRoot = Join-Path -Path $setupPath -ChildPath "Trash"
    $resolvedImportMode = Read-AsiToPixImportMode -ImportMode $ImportMode

    return [PSCustomObject]@{
        SourcePath     = $resolvedSourcePath
        AstroPhotoRoot = $resolvedAstroPhotoRoot
        ParsedFiles    = @($parsedFiles)
        ObjectName     = $objectSelection.Name
        SeasonName     = $resolvedSeasonName
        SetupName      = $resolvedSetupName
        ImportMode     = $resolvedImportMode
        SetupPath      = $setupPath
        GoodRoot       = $goodRoot
        TrashRoot      = $trashRoot
        AsiairRoot     = $asiairRoot
        ObjectPath     = $objectPath
        SeasonPath     = $seasonPath
    }
}

function Show-AsiToPixImportPlan {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Plan
    )

    Write-Host "`nImport target:" -ForegroundColor Cyan
    Write-Host "  Object:    $($Plan.ObjectName)" -ForegroundColor White
    Write-Host "  Season:    $($Plan.SeasonName)" -ForegroundColor White
    Write-Host "  Setup:     $($Plan.SetupName)" -ForegroundColor White
    Write-Host "  Mode:      $($Plan.ImportMode)" -ForegroundColor White
    Write-Host "  Root:      $($Plan.SetupPath)" -ForegroundColor White
    Write-Host "`nDetected groups:" -ForegroundColor Cyan
    $summaryItems = foreach ($entry in $Plan.ParsedFiles) {
        [PSCustomObject]@{
            FilterName             = $entry.FilterName
            DestinationNightFolder = Get-AsiToPixDestinationNightFolder -Entry $entry
        }
    }
    $summaryByFilterAndNight = $summaryItems |
        Group-Object -Property FilterName, DestinationNightFolder |
        Sort-Object Name

    foreach ($group in $summaryByFilterAndNight) {
        $first = $group.Group[0]
        Write-Host "  $($first.FilterName) / $($first.DestinationNightFolder): $($group.Count) file(s)" -ForegroundColor White
    }
}

function Invoke-AsiToPixImportPlan {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Plan
    )

    New-AsiToPixDirectory -Path $Plan.AsiairRoot -Cmdlet $PSCmdlet
    New-AsiToPixDirectory -Path $Plan.ObjectPath -Cmdlet $PSCmdlet
    New-AsiToPixDirectory -Path $Plan.SeasonPath -Cmdlet $PSCmdlet
    New-AsiToPixDirectory -Path $Plan.SetupPath -Cmdlet $PSCmdlet
    New-AsiToPixDirectory -Path $Plan.GoodRoot -Cmdlet $PSCmdlet
    New-AsiToPixDirectory -Path $Plan.TrashRoot -Cmdlet $PSCmdlet

    $goodNames = Get-AsiToPixFileNameSet -RootPath $Plan.GoodRoot
    $trashNames = Get-AsiToPixFileNameSet -RootPath $Plan.TrashRoot

    $imported = 0
    $skippedExisting = 0
    $skippedTrash = 0

    foreach ($entry in $Plan.ParsedFiles) {
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

        $filterGoodPath = Join-Path -Path $Plan.GoodRoot -ChildPath $entry.FilterName
        $filterTrashPath = Join-Path -Path $Plan.TrashRoot -ChildPath $entry.FilterName
        $destinationNightFolder = Get-AsiToPixDestinationNightFolder -Entry $entry
        $nightGoodPath = Join-Path -Path $filterGoodPath -ChildPath $destinationNightFolder
        $nightTrashPath = Join-Path -Path $filterTrashPath -ChildPath $destinationNightFolder
        $destinationFile = Join-Path -Path $nightGoodPath -ChildPath $fileName

        New-AsiToPixDirectory -Path $filterGoodPath -Cmdlet $PSCmdlet
        New-AsiToPixDirectory -Path $filterTrashPath -Cmdlet $PSCmdlet
        New-AsiToPixDirectory -Path $nightGoodPath -Cmdlet $PSCmdlet
        New-AsiToPixDirectory -Path $nightTrashPath -Cmdlet $PSCmdlet

        if (Test-Path -LiteralPath $destinationFile) {
            throw "Destination path already exists and will not be overwritten: $destinationFile"
        }

        $operationName = if ($Plan.ImportMode -eq "Symlink") { "Create symbolic link to '$($entry.File.FullName)'" } else { "Copy light file from '$($entry.File.FullName)'" }
        if ($PSCmdlet.ShouldProcess($destinationFile, $operationName)) {
            if ($Plan.ImportMode -eq "Symlink") {
                try {
                    New-Item -ItemType SymbolicLink -Path $destinationFile -Value $entry.File.FullName -ErrorAction Stop | Out-Null
                } catch {
                    throw "Failed to create symbolic link '$destinationFile' -> '$($entry.File.FullName)': $($_.Exception.Message)"
                }
            } else {
                try {
                    Copy-Item -LiteralPath $entry.File.FullName -Destination $destinationFile -ErrorAction Stop
                } catch {
                    throw "Failed to copy '$($entry.File.FullName)' to '$destinationFile': $($_.Exception.Message)"
                }
            }

            [void]$goodNames.Add($fileName)
            $imported++
            $displayOperation = if ($Plan.ImportMode -eq "Symlink") { "link" } else { "+" }
            Write-Host "  [$displayOperation] $fileName -> $($entry.FilterName)\$destinationNightFolder" -ForegroundColor Green
        }
    }

    return [PSCustomObject]@{
        Imported        = $imported
        AlreadyInGood   = $skippedExisting
        PreservedTrash  = $skippedTrash
        ImportMode      = $Plan.ImportMode
        SourcePath      = $Plan.SourcePath
        DestinationRoot = $Plan.SetupPath
    }
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

        [string]$CameraName = "",

        [ValidateSet("", "Copy", "Symlink")]
        [string]$ImportMode = ""
    )

    $plan = Get-AsiToPixImportPlan `
        -SourcePath $SourcePath `
        -AstroPhotoRoot $AstroPhotoRoot `
        -ObjectName $ObjectName `
        -SeasonName $SeasonName `
        -TelescopeName $TelescopeName `
        -CameraName $CameraName `
        -ImportMode $ImportMode

    Show-AsiToPixImportPlan -Plan $plan

    if (-not (Read-AsiToPixConfirmation -Prompt "Apply this import plan to the target tree?")) {
        Write-Host "[INFO] Import cancelled." -ForegroundColor Yellow
        return
    }

    $result = Invoke-AsiToPixImportPlan -Plan $plan -WhatIf:$WhatIfPreference

    Write-Host "`n[DONE] Import finished." -ForegroundColor Cyan
    if ($plan.ImportMode -eq "Symlink") {
        Write-Host "  Linked:           $($result.Imported)" -ForegroundColor White
    } else {
        Write-Host "  Copied:           $($result.Imported)" -ForegroundColor White
    }
    Write-Host "  Already in Good:  $($result.AlreadyInGood)" -ForegroundColor White
    Write-Host "  Preserved Trash:  $($result.PreservedTrash)" -ForegroundColor White
}

Export-ModuleMember -Function `
    ConvertTo-AsiToPixExposureFolderSuffix, `
    ConvertTo-AsiToPixPathSegment, `
    Find-AsiToPixImportSession, `
    Get-AsiToPixDefaultImportRoot, `
    Get-AsiToPixLightFileInfo, `
    Get-AsiToPixNameMatch, `
    Get-AsiToPixNightDate, `
    Get-AsiToPixSetupInfo, `
    Invoke-AsiToPixImportPlan, `
    Get-AsiToPixImportPlan, `
    Read-AsiToPixImportMode, `
    Resolve-AsiToPixImportSourcePath, `
    Resolve-AsiToPixLightFileInfo, `
    Show-AsiToPixImportPlan, `
    Import-AsiToPixSession
