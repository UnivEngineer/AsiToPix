Set-StrictMode -Version Latest

$imageFilesModule = Join-Path -Path $PSScriptRoot -ChildPath "AsiToPix.ImageFiles.psm1"
Import-Module $imageFilesModule -Force

function Read-AsiToPixCalibrationValue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Prompt,

        [string]$DefaultValue = "",

        [switch]$AllowEmpty
    )

    do {
        $displayPrompt = if ([string]::IsNullOrWhiteSpace($DefaultValue)) {
            $Prompt
        } else {
            "$Prompt [$DefaultValue]"
        }
        $value = (Read-Host $displayPrompt).Trim()
        if ([string]::IsNullOrWhiteSpace($value) -and -not [string]::IsNullOrWhiteSpace($DefaultValue)) {
            return $DefaultValue
        }
        if (-not [string]::IsNullOrWhiteSpace($value) -or $AllowEmpty) {
            return $value
        }

        Write-Host "[!] Value cannot be empty." -ForegroundColor Red
    } while ($true)
}

function Read-AsiToPixCalibrationConfirmation {
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

        $firstCharacter = $answer[0]
        if ($firstCharacter -in @([char]'y', [char]'Y', [char]0x0434, [char]0x0414)) { return $true }
        if ($firstCharacter -in @([char]'n', [char]'N', [char]0x043d, [char]0x041d)) { return $false }

        Write-Host "[!] Enter Y/N or the Cyrillic yes/no initials." -ForegroundColor Red
    } while ($true)
}

function ConvertTo-AsiToPixCalibrationNumericText {
    param(
        [AllowEmptyString()]
        [string]$Value,

        [string]$ValueName = "numeric value",

        [string]$UnitPattern = ""
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    $numericText = $Value.Trim().Replace(',', '.')
    if (-not [string]::IsNullOrWhiteSpace($UnitPattern)) {
        $numericText = $numericText -replace "(?i)$UnitPattern$", ""
    }

    $number = [decimal]0
    $parsed = [decimal]::TryParse(
        $numericText,
        [System.Globalization.NumberStyles]::Float,
        [System.Globalization.CultureInfo]::InvariantCulture,
        [ref]$number
    )
    if (-not $parsed) {
        throw "Invalid $ValueName '$Value'. Enter a number using a dot or comma as the decimal separator."
    }

    return $number.ToString("G29", [System.Globalization.CultureInfo]::InvariantCulture)
}

function Assert-AsiToPixCalibrationPathSegment {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Value,

        [Parameter(Mandatory = $true)]
        [string]$ValueName
    )

    if ([string]::IsNullOrWhiteSpace($Value) -or
        $Value -in @(".", "..") -or
        $Value.IndexOfAny([System.IO.Path]::GetInvalidFileNameChars()) -ge 0) {
        throw "Invalid $ValueName path segment: '$Value'."
    }
}

function Test-AsiToPixSupportedCalibrationFileName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileName
    )

    return (Test-AsiToPixSupportedImageFileName -FileName $FileName)
}

function Get-AsiToPixCalibrationCategoryName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FolderName
    )

    switch -Regex ($FolderName) {
        '^(?i:bias|biases)$' { return "Bias" }
        '^(?i:dark|darks)$' { return "Dark" }
        '^(?i:flat|flats)$' { return "Flat" }
        default { return $null }
    }
}

function ConvertFrom-AsiToPixCalibrationFileName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileName,

        [Parameter(Mandatory = $true)]
        [ValidateSet("Bias", "Dark", "Flat")]
        [string]$Category
    )

    $stem = Get-AsiToPixImageFileStem -FileName $FileName
    $exposureSeconds = $null
    if ($stem -match '^(?:Bias|Dark|Flat)_(?<value>\d+(?:\.\d+)?)(?<unit>ms|s)(?:_|$)') {
        $exposure = [decimal]::Parse(
            $Matches["value"],
            [System.Globalization.CultureInfo]::InvariantCulture
        )
        if ($Matches["unit"] -ieq "ms") {
            $exposure = $exposure / [decimal]1000
        }
        $exposureSeconds = $exposure.ToString("G29", [System.Globalization.CultureInfo]::InvariantCulture)
    }

    $cameraName = $null
    $filterName = $null
    if ($stem -match '_(?<camera>(?:ASI)?\d{3,4}M[MCP]?)(?:_(?<filter>[^_]+))?_gain') {
        $cameraName = $Matches["camera"]
        if ($cameraName -notmatch '^(?i:ASI)') {
            $cameraName = "ASI$cameraName"
        }
        if ($Matches.ContainsKey("filter") -and -not [string]::IsNullOrWhiteSpace($Matches["filter"])) {
            $filterName = $Matches["filter"].Trim()
        }
    }

    $gain = $null
    if ($stem -match '_(?i:gain)(?<gain>-?\d+(?:\.\d+)?)(?:_|$)') {
        $gain = ConvertTo-AsiToPixCalibrationNumericText -Value $Matches["gain"] -ValueName "gain"
    }

    $capturedAt = $null
    if ($stem -match '_(?<stamp>\d{8}-\d{6})(?:_|$)') {
        $capturedAt = [datetime]::ParseExact(
            $Matches["stamp"],
            "yyyyMMdd-HHmmss",
            [System.Globalization.CultureInfo]::InvariantCulture
        )
    }

    $temperatureC = $null
    if ($stem -match '_(?<temperature>-?\d+(?:\.\d+)?)C(?:_|$)') {
        $temperatureC = ConvertTo-AsiToPixCalibrationNumericText `
            -Value $Matches["temperature"] `
            -ValueName "temperature"
    }

    $angleDegrees = $null
    if ($stem -match '_(?<angle>-?\d+(?:\.\d+)?)deg(?:_|$)') {
        $angleDegrees = ConvertTo-AsiToPixCalibrationNumericText -Value $Matches["angle"] -ValueName "angle"
    }

    return [PSCustomObject]@{
        Category        = $Category
        CameraName      = $cameraName
        Gain            = $gain
        TemperatureC    = $temperatureC
        ExposureSeconds = $exposureSeconds
        FilterName      = $filterName
        AngleDegrees    = $angleDegrees
        CapturedAt      = $capturedAt
    }
}

function Get-AsiToPixCalibrationSourceRecord {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath
    )

    $resolvedSourcePath = (Resolve-Path -LiteralPath $SourcePath).ProviderPath
    $categoryFolders = @(Get-ChildItem -LiteralPath $resolvedSourcePath -Directory -ErrorAction Stop |
        ForEach-Object {
            $category = Get-AsiToPixCalibrationCategoryName -FolderName $_.Name
            if ($null -ne $category) {
                [PSCustomObject]@{
                    Directory = $_
                    Category  = $category
                }
            }
        })

    if ($categoryFolders.Count -eq 0) {
        throw "No flat(s), dark(s), or bias(es) folders found directly under import root: $resolvedSourcePath"
    }

    $records = @()
    foreach ($categoryFolder in $categoryFolders) {
        $files = @(Get-ChildItem -LiteralPath $categoryFolder.Directory.FullName -File -Recurse -ErrorAction Stop |
            Where-Object { Test-AsiToPixSupportedCalibrationFileName -FileName $_.Name } |
            Sort-Object FullName)

        foreach ($file in $files) {
            $metadata = ConvertFrom-AsiToPixCalibrationFileName `
                -FileName $file.Name `
                -Category $categoryFolder.Category
            $capturedAt = if ($null -ne $metadata.CapturedAt) { $metadata.CapturedAt } else { $file.LastWriteTime }
            $records += [PSCustomObject]@{
                File             = $file
                Category         = $metadata.Category
                CameraName       = $metadata.CameraName
                Gain             = $metadata.Gain
                TemperatureC     = $metadata.TemperatureC
                ExposureSeconds  = $metadata.ExposureSeconds
                FilterName       = $metadata.FilterName
                AngleDegrees     = $metadata.AngleDegrees
                CapturedAt       = $capturedAt
                UsedFileTime     = ($null -eq $metadata.CapturedAt)
            }
        }
    }

    if ($records.Count -eq 0) {
        throw "No supported PixInsight calibration image files found under import root: $resolvedSourcePath"
    }

    return @($records)
}

function ConvertTo-AsiToPixCalibrationFilterName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FilterName,

        [Parameter(Mandatory = $true)]
        [string]$CameraName
    )

    $filter = $FilterName.Trim()
    if ($CameraName -match '(?i)MM(?:PRO)?$') {
        switch -Regex ($filter) {
            '^(?i:H|Ha)$' { return "H" }
            '^(?i:S|SII)$' { return "S" }
            '^(?i:O|OII|OIII)$' { return "O" }
            '^(?i:L|R|G|B|None)$' { return $filter.ToUpperInvariant().Replace("NONE", "None") }
            default { return $filter }
        }
    }

    switch -Regex ($filter) {
        '^(?i:IRC|Trib|L)$' { return "L" }
        '^(?i:None)$' { return "None" }
        '^(?i:HO|UHC)$' { return $filter.ToUpperInvariant() }
        '^(?i:SO)$' { return "SO" }
        default { return $filter }
    }
}

function Get-AsiToPixCalibrationTemperatureFolder {
    param(
        [Parameter(Mandatory = $true)]
        [string]$TemperatureC
    )

    $canonicalTemperature = ConvertTo-AsiToPixCalibrationNumericText `
        -Value $TemperatureC `
        -ValueName "temperature" `
        -UnitPattern 'C'
    $temperature = [decimal]::Parse(
        $canonicalTemperature,
        [System.Globalization.CultureInfo]::InvariantCulture
    )
    $roundedTemperature = [Math]::Round($temperature / [decimal]5) * [decimal]5
    if ($roundedTemperature -eq [decimal]0) {
        $roundedTemperature = [decimal]0
    }

    return "$($roundedTemperature.ToString('G29', [System.Globalization.CultureInfo]::InvariantCulture))C"
}

function Get-AsiToPixCalibrationNightStart {
    param(
        [Parameter(Mandatory = $true)]
        [datetime]$CapturedAt
    )

    $nightStart = $CapturedAt.Date
    if ($CapturedAt.Hour -lt 12) {
        $nightStart = $nightStart.AddDays(-1)
    }

    return $nightStart
}

function Get-AsiToPixCalibrationDestinationFolder {
    param(
        [Parameter(Mandatory = $true)]
        [object]$SourceRecord,

        [Parameter(Mandatory = $true)]
        [string]$CalibrationRoot,

        [Parameter(Mandatory = $true)]
        [string]$SetupName,

        [string]$CameraName = "",

        [string]$Gain = "",

        [string]$TemperatureC = "",

        [string]$DarkExposureSeconds = "",

        [string]$FilterName = "",

        [string]$AngleDegrees = ""
    )

    $resolvedCameraName = if (-not [string]::IsNullOrWhiteSpace($SourceRecord.CameraName)) {
        [string]$SourceRecord.CameraName
    } else {
        $CameraName
    }
    if ([string]::IsNullOrWhiteSpace($resolvedCameraName)) {
        throw "Camera name is missing for calibration file: $($SourceRecord.File.FullName)"
    }
    Assert-AsiToPixCalibrationPathSegment -Value $resolvedCameraName -ValueName "camera name"

    $cameraSourceRoot = Join-Path -Path $CalibrationRoot -ChildPath $resolvedCameraName
    $cameraSourceRoot = Join-Path -Path $cameraSourceRoot -ChildPath "Source"
    $nightStart = Get-AsiToPixCalibrationNightStart -CapturedAt $SourceRecord.CapturedAt
    $monthName = $nightStart.ToString(
        "yy.MM",
        [System.Globalization.CultureInfo]::InvariantCulture
    )

    if ($SourceRecord.Category -eq "Flat") {
        Assert-AsiToPixCalibrationPathSegment -Value $SetupName -ValueName "setup name"
        $resolvedFilterName = if (-not [string]::IsNullOrWhiteSpace($SourceRecord.FilterName)) {
            [string]$SourceRecord.FilterName
        } else {
            $FilterName
        }
        if ([string]::IsNullOrWhiteSpace($resolvedFilterName)) {
            throw "Filter name is missing for flat file: $($SourceRecord.File.FullName)"
        }
        $resolvedFilterName = ConvertTo-AsiToPixCalibrationFilterName `
            -FilterName $resolvedFilterName `
            -CameraName $resolvedCameraName
        Assert-AsiToPixCalibrationPathSegment -Value $resolvedFilterName -ValueName "filter name"

        $dateName = $nightStart.ToString(
            "yy.MM.dd",
            [System.Globalization.CultureInfo]::InvariantCulture
        )
        $resolvedAngle = if (-not [string]::IsNullOrWhiteSpace($SourceRecord.AngleDegrees)) {
            [string]$SourceRecord.AngleDegrees
        } else {
            $AngleDegrees
        }
        $folderName = "$dateName $resolvedFilterName"
        if (-not [string]::IsNullOrWhiteSpace($resolvedAngle)) {
            $canonicalAngle = ConvertTo-AsiToPixCalibrationNumericText `
                -Value $resolvedAngle `
                -ValueName "flat angle" `
                -UnitPattern 'deg'
            $folderName = "$folderName ${canonicalAngle}deg"
        }

        $flatRoot = Join-Path -Path $cameraSourceRoot -ChildPath "flats"
        $setupRoot = Join-Path -Path $flatRoot -ChildPath $SetupName
        return (Join-Path -Path $setupRoot -ChildPath $folderName)
    }

    $resolvedGain = if (-not [string]::IsNullOrWhiteSpace($SourceRecord.Gain)) {
        [string]$SourceRecord.Gain
    } else {
        $Gain
    }
    $resolvedTemperature = if (-not [string]::IsNullOrWhiteSpace($SourceRecord.TemperatureC)) {
        [string]$SourceRecord.TemperatureC
    } else {
        $TemperatureC
    }
    if ([string]::IsNullOrWhiteSpace($resolvedGain)) {
        throw "Gain/ISO is missing for calibration file: $($SourceRecord.File.FullName)"
    }
    if ([string]::IsNullOrWhiteSpace($resolvedTemperature)) {
        throw "Temperature is missing for calibration file: $($SourceRecord.File.FullName)"
    }

    $canonicalGain = ConvertTo-AsiToPixCalibrationNumericText -Value $resolvedGain -ValueName "gain/ISO"
    $temperatureFolder = Get-AsiToPixCalibrationTemperatureFolder -TemperatureC $resolvedTemperature
    $kindName = if ($SourceRecord.Category -eq "Bias") { "biases" } else { "darks" }
    $kindRoot = Join-Path -Path $cameraSourceRoot -ChildPath $kindName
    $gainRoot = Join-Path -Path $kindRoot -ChildPath "gain$canonicalGain"
    $temperatureRoot = Join-Path -Path $gainRoot -ChildPath $temperatureFolder

    if ($SourceRecord.Category -eq "Bias") {
        return (Join-Path -Path $temperatureRoot -ChildPath $monthName)
    }

    $resolvedExposure = if (-not [string]::IsNullOrWhiteSpace($SourceRecord.ExposureSeconds)) {
        [string]$SourceRecord.ExposureSeconds
    } else {
        $DarkExposureSeconds
    }
    if ([string]::IsNullOrWhiteSpace($resolvedExposure)) {
        throw "Exposure is missing for dark file: $($SourceRecord.File.FullName)"
    }
    $canonicalExposure = ConvertTo-AsiToPixCalibrationNumericText `
        -Value $resolvedExposure `
        -ValueName "dark exposure" `
        -UnitPattern '(?:s|sec)'
    $exposureRoot = Join-Path -Path $temperatureRoot -ChildPath "${canonicalExposure}sec"
    return (Join-Path -Path $exposureRoot -ChildPath $monthName)
}

function ConvertTo-AsiToPixCalibrationImportPlan {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$SourceRecord,

        [Parameter(Mandatory = $true)]
        [string]$CalibrationRoot,

        [Parameter(Mandatory = $true)]
        [string]$SetupName,

        [string]$CameraName = "",

        [string]$Gain = "",

        [string]$TemperatureC = "",

        [string]$DarkExposureSeconds = "",

        [string]$FilterName = "",

        [string]$AngleDegrees = ""
    )

    if (-not (Test-Path -LiteralPath $CalibrationRoot -PathType Container)) {
        throw "Calibration root not found: $CalibrationRoot"
    }
    $resolvedCalibrationRoot = (Resolve-Path -LiteralPath $CalibrationRoot).ProviderPath
    Assert-AsiToPixCalibrationPathSegment -Value $SetupName -ValueName "setup name"

    $entries = @()
    $plannedPaths = @{}
    foreach ($record in $SourceRecord) {
        $destinationFolder = Get-AsiToPixCalibrationDestinationFolder `
            -SourceRecord $record `
            -CalibrationRoot $resolvedCalibrationRoot `
            -SetupName $SetupName `
            -CameraName $CameraName `
            -Gain $Gain `
            -TemperatureC $TemperatureC `
            -DarkExposureSeconds $DarkExposureSeconds `
            -FilterName $FilterName `
            -AngleDegrees $AngleDegrees
        $destinationPath = Join-Path -Path $destinationFolder -ChildPath $record.File.Name
        $status = "Planned"
        $reason = ""
        $existingPath = $null

        if (Test-Path -LiteralPath $destinationPath) {
            if (-not (Test-Path -LiteralPath $destinationPath -PathType Leaf)) {
                $status = "Conflict"
                $reason = "Destination path exists but is not an ordinary file."
            } else {
                $existingFile = Get-Item -LiteralPath $destinationPath -Force -ErrorAction Stop
                $existingPath = $existingFile.FullName
                if ($existingFile.Length -eq $record.File.Length) {
                    $status = "Exists"
                    $reason = "A file with the same name and size already exists."
                } else {
                    $status = "Conflict"
                    $reason = "A file with the same name but a different size already exists."
                }
            }
        } elseif ($plannedPaths.ContainsKey($destinationPath)) {
            $status = "Conflict"
            $reason = "More than one source file resolves to the same destination path."
        } else {
            $plannedPaths[$destinationPath] = $record.File.FullName
        }

        $entries += [PSCustomObject]@{
            Status            = $status
            Reason            = $reason
            Category          = $record.Category
            SourcePath        = $record.File.FullName
            SourceLength      = $record.File.Length
            DestinationFolder = $destinationFolder
            DestinationPath   = $destinationPath
            ExistingPath      = $existingPath
            UsedFileTime      = $record.UsedFileTime
        }
    }

    $additionWarnings = @()
    foreach ($group in @($entries | Where-Object { $_.Status -eq "Planned" } | Group-Object DestinationFolder)) {
        $existingCount = 0
        if (Test-Path -LiteralPath $group.Name -PathType Container) {
            $existingCount = @(Get-ChildItem -LiteralPath $group.Name -File -ErrorAction Stop).Count
        }
        if ($existingCount -gt 0) {
            $additionWarnings += [PSCustomObject]@{
                DestinationFolder = $group.Name
                ExistingCount     = $existingCount
                NewCount          = $group.Count
            }
        }
    }

    return [PSCustomObject]@{
        SourceCount      = $SourceRecord.Count
        PlannedCount     = @($entries | Where-Object { $_.Status -eq "Planned" }).Count
        ExistingCount    = @($entries | Where-Object { $_.Status -eq "Exists" }).Count
        ConflictCount    = @($entries | Where-Object { $_.Status -eq "Conflict" }).Count
        Entries          = @($entries)
        AdditionWarnings = @($additionWarnings)
    }
}

function Get-AsiToPixCalibrationImportPlan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,

        [Parameter(Mandatory = $true)]
        [string]$CalibrationRoot,

        [string]$SetupName = "",

        [string]$CameraName = "",

        [string]$Gain = "",

        [string]$TemperatureC = "",

        [string]$DarkExposureSeconds = "",

        [string]$FilterName = "",

        [string]$AngleDegrees = ""
    )

    $resolvedSourcePath = (Resolve-Path -LiteralPath $SourcePath).ProviderPath
    if ([string]::IsNullOrWhiteSpace($SetupName)) {
        $SetupName = Split-Path -Path $resolvedSourcePath -Leaf
    }
    $sourceRecords = @(Get-AsiToPixCalibrationSourceRecord -SourcePath $resolvedSourcePath)
    return ConvertTo-AsiToPixCalibrationImportPlan `
        -SourceRecord $sourceRecords `
        -CalibrationRoot $CalibrationRoot `
        -SetupName $SetupName `
        -CameraName $CameraName `
        -Gain $Gain `
        -TemperatureC $TemperatureC `
        -DarkExposureSeconds $DarkExposureSeconds `
        -FilterName $FilterName `
        -AngleDegrees $AngleDegrees
}

function Get-AsiToPixCalibrationDefaultValue {
    param(
        [Parameter(Mandatory = $true)]
        [AllowNull()]
        [AllowEmptyCollection()]
        [object[]]$Value
    )

    $values = @($Value | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Select-Object -Unique)
    if ($values.Count -eq 1) {
        return [string]$values[0]
    }

    return ""
}

function Write-AsiToPixCalibrationMissingMetadataExample {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Description,

        [Parameter(Mandatory = $true)]
        [string]$SourcePath,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$Record,

        [int]$MaximumExampleCount = 3
    )

    if ($Record.Count -eq 0) {
        return
    }

    $separatorCharacters = [char[]]@(
        [System.IO.Path]::DirectorySeparatorChar,
        [System.IO.Path]::AltDirectorySeparatorChar
    )
    $normalizedSourcePath = [System.IO.Path]::GetFullPath($SourcePath).TrimEnd($separatorCharacters)
    $sourcePrefix = "$normalizedSourcePath$([System.IO.Path]::DirectorySeparatorChar)"

    Write-Host "`n[INFO] $Description is missing for $($Record.Count) file(s). Examples:" -ForegroundColor Yellow
    foreach ($example in @($Record | Select-Object -First $MaximumExampleCount)) {
        $filePath = [System.IO.Path]::GetFullPath($example.File.FullName)
        $displayPath = if ($filePath.StartsWith($sourcePrefix, [System.StringComparison]::OrdinalIgnoreCase)) {
            $filePath.Substring($sourcePrefix.Length)
        } else {
            $filePath
        }
        Write-Host "  $displayPath" -ForegroundColor DarkGray
    }

    $remainingCount = $Record.Count - [Math]::Min($Record.Count, $MaximumExampleCount)
    if ($remainingCount -gt 0) {
        Write-Host "  ... and $remainingCount more file(s)" -ForegroundColor DarkGray
    }
}

function Read-AsiToPixCalibrationCameraName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$CalibrationRoot,

        [string]$DefaultValue = ""
    )

    if (-not [string]::IsNullOrWhiteSpace($DefaultValue)) {
        return Read-AsiToPixCalibrationValue -Prompt "Enter camera name" -DefaultValue $DefaultValue
    }

    $cameras = @(Get-ChildItem -LiteralPath $CalibrationRoot -Directory -ErrorAction Stop |
        Sort-Object Name |
        Select-Object -ExpandProperty Name)
    if ($cameras.Count -gt 0) {
        Write-Host "`nExisting calibration cameras:" -ForegroundColor Cyan
        for ($index = 0; $index -lt $cameras.Count; $index++) {
            Write-Host " [$($index + 1)] $($cameras[$index])" -ForegroundColor White
        }
        Write-Host " [0] Enter a new camera name" -ForegroundColor White

        do {
            $answer = (Read-Host "Select camera index or type a camera name").Trim()
            $selectedIndex = -1
            if ([int]::TryParse($answer, [ref]$selectedIndex)) {
                if ($selectedIndex -eq 0) { break }
                if ($selectedIndex -ge 1 -and $selectedIndex -le $cameras.Count) {
                    return $cameras[$selectedIndex - 1]
                }
            } elseif (-not [string]::IsNullOrWhiteSpace($answer)) {
                return $answer
            }

            Write-Host "[!] Invalid selection." -ForegroundColor Red
        } while ($true)
    }

    return Read-AsiToPixCalibrationValue -Prompt "Enter camera name"
}

function Write-AsiToPixCalibrationImportPlan {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Plan
    )

    Write-Host "`nImport plan:" -ForegroundColor Cyan
    foreach ($entry in $Plan.Entries) {
        $fileName = Split-Path -Path $entry.SourcePath -Leaf
        switch ($entry.Status) {
            "Exists" {
                Write-Host "  [exists] $fileName -> $($entry.DestinationFolder)" -ForegroundColor DarkGray
            }
            "Conflict" {
                Write-Host "  [conflict] $fileName -> $($entry.DestinationFolder)" -ForegroundColor Red
                Write-Host "    $($entry.Reason)" -ForegroundColor Red
            }
            default {
                Write-Host "  [+] $fileName -> $($entry.DestinationFolder)" -ForegroundColor Green
            }
        }
    }

    foreach ($warning in $Plan.AdditionWarnings) {
        Write-Host "`n[!] Existing destination will receive new files (unusual case):" -ForegroundColor Yellow
        Write-Host "    $($warning.DestinationFolder)" -ForegroundColor Yellow
        Write-Host "    Existing: $($warning.ExistingCount); new: $($warning.NewCount)" -ForegroundColor DarkYellow
    }

    Write-Host "`nSummary:" -ForegroundColor Cyan
    Write-Host "  New:       $($Plan.PlannedCount)" -ForegroundColor White
    Write-Host "  Existing:  $($Plan.ExistingCount)" -ForegroundColor White
    Write-Host "  Conflicts: $($Plan.ConflictCount)" -ForegroundColor White
}

function Invoke-AsiToPixCalibrationImportPlan {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Plan
    )

    $copiedCount = 0
    $whatIfCount = 0
    foreach ($entry in @($Plan.Entries | Where-Object { $_.Status -eq "Planned" })) {
        if (-not $PSCmdlet.ShouldProcess(
            $entry.DestinationPath,
            "Copy calibration file from '$($entry.SourcePath)'"
        )) {
            $whatIfCount++
            continue
        }

        try {
            if (Test-Path -LiteralPath $entry.DestinationFolder) {
                if (-not (Test-Path -LiteralPath $entry.DestinationFolder -PathType Container)) {
                    throw "Destination folder path is occupied by a non-directory item."
                }
            } else {
                New-Item -ItemType Directory -Path $entry.DestinationFolder -Force -ErrorAction Stop | Out-Null
            }

            if (Test-Path -LiteralPath $entry.DestinationPath) {
                throw "Destination file appeared after planning; refusing to overwrite it."
            }

            Copy-Item `
                -LiteralPath $entry.SourcePath `
                -Destination $entry.DestinationPath `
                -ErrorAction Stop
            $copiedCount++
            Write-Host "  [+] $(Split-Path -Path $entry.SourcePath -Leaf) -> $($entry.DestinationFolder)" `
                -ForegroundColor Green
        } catch {
            throw "Failed to copy calibration file '$($entry.SourcePath)' to '$($entry.DestinationPath)': $($_.Exception.Message)"
        }
    }

    return [PSCustomObject]@{
        CopiedCount   = $copiedCount
        ExistingCount = $Plan.ExistingCount
        ConflictCount = $Plan.ConflictCount
        WhatIfCount   = $whatIfCount
    }
}

function Import-AsiToPixCalibration {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,

        [Parameter(Mandatory = $true)]
        [string]$CalibrationRoot,

        [string]$CameraName = "",

        [string]$Gain = "",

        [string]$TemperatureC = "",

        [string]$DarkExposureSeconds = "",

        [string]$FilterName = "",

        [string]$AngleDegrees = ""
    )

    $resolvedSourcePath = (Resolve-Path -LiteralPath $SourcePath).ProviderPath
    $resolvedCalibrationRoot = (Resolve-Path -LiteralPath $CalibrationRoot).ProviderPath
    $setupName = Split-Path -Path $resolvedSourcePath -Leaf
    $records = @(Get-AsiToPixCalibrationSourceRecord -SourcePath $resolvedSourcePath)

    Write-Host "`nSource setup: $setupName" -ForegroundColor Cyan
    Write-Host "Found $($records.Count) calibration file(s)." -ForegroundColor White

    $missingCamera = @($records | Where-Object { [string]::IsNullOrWhiteSpace($_.CameraName) })
    if ($missingCamera.Count -gt 0 -and [string]::IsNullOrWhiteSpace($CameraName)) {
        $detectedCamera = Get-AsiToPixCalibrationDefaultValue -Value @($records.CameraName)
        $CameraName = Read-AsiToPixCalibrationCameraName `
            -CalibrationRoot $resolvedCalibrationRoot `
            -DefaultValue $detectedCamera
    }

    $biasAndDark = @($records | Where-Object { $_.Category -in @("Bias", "Dark") })
    if (@($biasAndDark | Where-Object { [string]::IsNullOrWhiteSpace($_.Gain) }).Count -gt 0 -and
        [string]::IsNullOrWhiteSpace($Gain)) {
        $detectedGain = Get-AsiToPixCalibrationDefaultValue -Value @($biasAndDark.Gain)
        $Gain = Read-AsiToPixCalibrationValue -Prompt "Enter gain/ISO for files without filename metadata" -DefaultValue $detectedGain
    }
    if (@($biasAndDark | Where-Object { [string]::IsNullOrWhiteSpace($_.TemperatureC) }).Count -gt 0 -and
        [string]::IsNullOrWhiteSpace($TemperatureC)) {
        $detectedTemperature = Get-AsiToPixCalibrationDefaultValue -Value @($biasAndDark.TemperatureC)
        $TemperatureC = Read-AsiToPixCalibrationValue `
            -Prompt "Enter temperature in C for files without filename metadata" `
            -DefaultValue $detectedTemperature
    }

    $darks = @($records | Where-Object { $_.Category -eq "Dark" })
    if (@($darks | Where-Object { [string]::IsNullOrWhiteSpace($_.ExposureSeconds) }).Count -gt 0 -and
        [string]::IsNullOrWhiteSpace($DarkExposureSeconds)) {
        $detectedExposure = Get-AsiToPixCalibrationDefaultValue -Value @($darks.ExposureSeconds)
        $DarkExposureSeconds = Read-AsiToPixCalibrationValue `
            -Prompt "Enter dark exposure in seconds for files without filename metadata" `
            -DefaultValue $detectedExposure
    }

    $flats = @($records | Where-Object { $_.Category -eq "Flat" })
    $flatsWithoutFilter = @($flats | Where-Object { [string]::IsNullOrWhiteSpace($_.FilterName) })
    if ($flatsWithoutFilter.Count -gt 0 -and
        [string]::IsNullOrWhiteSpace($FilterName)) {
        Write-AsiToPixCalibrationMissingMetadataExample `
            -Description "Flat filter metadata" `
            -SourcePath $resolvedSourcePath `
            -Record $flatsWithoutFilter
        $detectedFilter = Get-AsiToPixCalibrationDefaultValue -Value @($flats.FilterName)
        if ([string]::IsNullOrWhiteSpace($detectedFilter)) { $detectedFilter = "None" }
        $FilterName = Read-AsiToPixCalibrationValue `
            -Prompt "Enter flat filter for files without filename metadata" `
            -DefaultValue $detectedFilter
    }
    if (@($flats | Where-Object { [string]::IsNullOrWhiteSpace($_.AngleDegrees) }).Count -gt 0 -and
        [string]::IsNullOrWhiteSpace($AngleDegrees)) {
        $detectedAngle = Get-AsiToPixCalibrationDefaultValue -Value @($flats.AngleDegrees)
        $AngleDegrees = Read-AsiToPixCalibrationValue `
            -Prompt "Enter flat angle in degrees, or press Enter to omit it" `
            -DefaultValue $detectedAngle `
            -AllowEmpty
    }

    $fallbackTimestampCount = @($records | Where-Object { $_.UsedFileTime }).Count
    if ($fallbackTimestampCount -gt 0) {
        Write-Host "[INFO] Using file timestamps for $fallbackTimestampCount file(s) without ASIAir timestamps." `
            -ForegroundColor DarkGray
    }

    $plan = ConvertTo-AsiToPixCalibrationImportPlan `
        -SourceRecord $records `
        -CalibrationRoot $resolvedCalibrationRoot `
        -SetupName $setupName `
        -CameraName $CameraName `
        -Gain $Gain `
        -TemperatureC $TemperatureC `
        -DarkExposureSeconds $DarkExposureSeconds `
        -FilterName $FilterName `
        -AngleDegrees $AngleDegrees
    Write-AsiToPixCalibrationImportPlan -Plan $plan

    if ($plan.PlannedCount -eq 0) {
        Write-Host "`n[DONE] No new calibration files to copy." -ForegroundColor Cyan
        return [PSCustomObject]@{
            CopiedCount    = 0
            ExistingCount  = $plan.ExistingCount
            ConflictCount  = $plan.ConflictCount
            WhatIfCount    = 0
            Cancelled      = $false
        }
    }

    if (-not $WhatIfPreference -and
        -not (Read-AsiToPixCalibrationConfirmation -Prompt "Copy $($plan.PlannedCount) new calibration file(s)?")) {
        Write-Host "[INFO] Calibration import cancelled." -ForegroundColor Yellow
        return [PSCustomObject]@{
            CopiedCount    = 0
            ExistingCount  = $plan.ExistingCount
            ConflictCount  = $plan.ConflictCount
            WhatIfCount    = 0
            Cancelled      = $true
        }
    }

    $result = Invoke-AsiToPixCalibrationImportPlan `
        -Plan $plan `
        -WhatIf:$WhatIfPreference `
        -Confirm:$false
    Write-Host "`n[DONE] Calibration import finished." -ForegroundColor Cyan
    Write-Host "  Copied:   $($result.CopiedCount)" -ForegroundColor White
    Write-Host "  Existing: $($result.ExistingCount)" -ForegroundColor White
    Write-Host "  Conflicts: $($result.ConflictCount)" -ForegroundColor White

    return $result
}

Export-ModuleMember -Function `
    ConvertFrom-AsiToPixCalibrationFileName, `
    Get-AsiToPixCalibrationImportPlan, `
    Import-AsiToPixCalibration, `
    Invoke-AsiToPixCalibrationImportPlan, `
    Test-AsiToPixSupportedCalibrationFileName
