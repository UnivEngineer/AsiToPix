Set-StrictMode -Version Latest

$imageFilesModule = Join-Path -Path $PSScriptRoot -ChildPath "AsiToPix.ImageFiles.psm1"
Import-Module $imageFilesModule -Force

$frameFoldersModule = Join-Path -Path $PSScriptRoot -ChildPath "AsiToPix.FrameFolders.psm1"
Import-Module $frameFoldersModule -Force

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

    foreach ($kind in @("Bias", "Dark", "Flat")) {
        if (Test-AsiToPixFrameFolderName -Name $FolderName -Kind $kind) {
            return $kind
        }
    }

    return $null
}

function Find-AsiToPixCalibrationImportFolder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ImportRoot,

        [ValidateSet("Bias", "Dark", "Flat")]
        [string[]]$Category = @("Bias", "Dark", "Flat")
    )

    if (-not (Test-Path -LiteralPath $ImportRoot -PathType Container)) {
        return @()
    }

    $resolvedImportRoot = (Resolve-Path -LiteralPath $ImportRoot).ProviderPath
    $folders = foreach ($setupFolder in Get-ChildItem -LiteralPath $resolvedImportRoot -Directory -ErrorAction Stop) {
        foreach ($categoryFolder in Get-ChildItem -LiteralPath $setupFolder.FullName -Directory -ErrorAction Stop) {
            $detectedCategory = Get-AsiToPixCalibrationCategoryName -FolderName $categoryFolder.Name
            if ($detectedCategory -in $Category) {
                [PSCustomObject]@{
                    SetupName  = $setupFolder.Name
                    SourcePath = $categoryFolder.FullName
                    Category   = $detectedCategory
                }
            }
        }
    }

    return @($folders | Sort-Object SetupName, SourcePath)
}

function Get-AsiToPixCalibrationSetupName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath
    )

    $resolvedSourcePath = (Resolve-Path -LiteralPath $SourcePath).ProviderPath
    $sourceName = Split-Path -Path $resolvedSourcePath -Leaf
    if ($null -ne (Get-AsiToPixCalibrationCategoryName -FolderName $sourceName)) {
        return (Split-Path -Path (Split-Path -Path $resolvedSourcePath -Parent) -Leaf)
    }

    return $sourceName
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
    $sourceDirectory = Get-Item -LiteralPath $resolvedSourcePath -Force -ErrorAction Stop
    $sourceCategory = Get-AsiToPixCalibrationCategoryName -FolderName $sourceDirectory.Name
    $categoryFolders = @(
        if ($null -ne $sourceCategory) {
            [PSCustomObject]@{
                Directory = $sourceDirectory
                Category  = $sourceCategory
            }
        } else {
            Get-ChildItem -LiteralPath $resolvedSourcePath -Directory -ErrorAction Stop | ForEach-Object {
                $category = Get-AsiToPixCalibrationCategoryName -FolderName $_.Name
                if ($null -ne $category) {
                    [PSCustomObject]@{
                        Directory = $_
                        Category  = $category
                    }
                }
            }
        }
    )

    if ($categoryFolders.Count -eq 0) {
        throw "Source is not a calibration category folder and contains no flat(s), dark(s), or bias(es) folders: $resolvedSourcePath"
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

function ConvertTo-AsiToPixFlatExposureFolderSuffix {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ExposureSeconds
    )

    $canonicalExposure = ConvertTo-AsiToPixCalibrationNumericText `
        -Value $ExposureSeconds `
        -ValueName "flat exposure"
    $exposure = [decimal]::Parse(
        $canonicalExposure,
        [System.Globalization.CultureInfo]::InvariantCulture
    )
    if ($exposure -lt [decimal]1) {
        $milliseconds = $exposure * [decimal]1000
        return "$($milliseconds.ToString('G29', [System.Globalization.CultureInfo]::InvariantCulture))ms"
    }

    return "${canonicalExposure}s"
}

function Get-AsiToPixFlatFolderExposure {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        return @()
    }

    $exposures = foreach ($file in Get-ChildItem -LiteralPath $Path -File -ErrorAction Stop) {
        if (-not (Test-AsiToPixSupportedCalibrationFileName -FileName $file.Name)) {
            continue
        }

        $metadata = ConvertFrom-AsiToPixCalibrationFileName -FileName $file.Name -Category Flat
        if (-not [string]::IsNullOrWhiteSpace($metadata.ExposureSeconds)) {
            ConvertTo-AsiToPixCalibrationNumericText `
                -Value ([string]$metadata.ExposureSeconds) `
                -ValueName "flat exposure"
        }
    }

    return @($exposures | Sort-Object -Unique)
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

        $flatFolderName = Get-AsiToPixCanonicalFrameFolderName -Kind Flat
        $flatRoot = Join-Path -Path $cameraSourceRoot -ChildPath $flatFolderName
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
    $kindName = if ($SourceRecord.Category -eq "Bias") {
        Get-AsiToPixCanonicalFrameFolderName -Kind Bias
    } else {
        Get-AsiToPixCanonicalFrameFolderName -Kind Dark
    }
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

    $destinationRecords = @(
        foreach ($record in $SourceRecord) {
            $baseDestinationFolder = Get-AsiToPixCalibrationDestinationFolder `
                -SourceRecord $record `
                -CalibrationRoot $resolvedCalibrationRoot `
                -SetupName $SetupName `
                -CameraName $CameraName `
                -Gain $Gain `
                -TemperatureC $TemperatureC `
                -DarkExposureSeconds $DarkExposureSeconds `
                -FilterName $FilterName `
                -AngleDegrees $AngleDegrees
            $flatExposureKey = if ($record.Category -eq "Flat" -and
                -not [string]::IsNullOrWhiteSpace($record.ExposureSeconds)) {
                ConvertTo-AsiToPixCalibrationNumericText `
                    -Value ([string]$record.ExposureSeconds) `
                    -ValueName "flat exposure"
            } else {
                $null
            }

            [PSCustomObject]@{
                Record                = $record
                BaseDestinationFolder = $baseDestinationFolder
                DestinationFolder     = $baseDestinationFolder
                FlatExposureKey       = $flatExposureKey
            }
        }
    )

    $flatExposureSplits = @()
    $reportedFlatSplits = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    foreach ($flatGroup in @(
        $destinationRecords |
            Where-Object { $_.Record.Category -eq "Flat" } |
            Group-Object BaseDestinationFolder
    )) {
        $baseFolder = $flatGroup.Name
        $baseImageFiles = @(
            if (Test-Path -LiteralPath $baseFolder -PathType Container) {
                Get-ChildItem -LiteralPath $baseFolder -File -ErrorAction Stop |
                    Where-Object { Test-AsiToPixSupportedCalibrationFileName -FileName $_.Name }
            }
        )
        $baseExposureKeys = [System.Collections.Generic.HashSet[string]]::new(
            [System.StringComparer]::OrdinalIgnoreCase
        )
        foreach ($existingExposure in @(Get-AsiToPixFlatFolderExposure -Path $baseFolder)) {
            [void]$baseExposureKeys.Add([string]$existingExposure)
        }

        $knownFlatRecords = @($flatGroup.Group | Where-Object {
            -not [string]::IsNullOrWhiteSpace($_.FlatExposureKey)
        })
        $unknownFlatRecords = @($flatGroup.Group | Where-Object {
            [string]::IsNullOrWhiteSpace($_.FlatExposureKey)
        })
        if ($baseImageFiles.Count -eq 0 -and
            $unknownFlatRecords.Count -eq 0 -and
            $knownFlatRecords.Count -gt 0) {
            $primaryFlatRecord = $knownFlatRecords |
                Sort-Object `
                    @{ Expression = { $_.Record.CapturedAt } }, `
                    @{ Expression = { $_.Record.File.FullName } } |
                Select-Object -First 1
            [void]$baseExposureKeys.Add([string]$primaryFlatRecord.FlatExposureKey)
        }

        foreach ($destinationRecord in $knownFlatRecords) {
            $exposureKey = [string]$destinationRecord.FlatExposureKey
            if ($baseExposureKeys.Contains($exposureKey)) {
                continue
            }

            $exposureSuffix = ConvertTo-AsiToPixFlatExposureFolderSuffix -ExposureSeconds $exposureKey
            $suffixFolder = "$baseFolder $exposureSuffix"
            $suffixImageFiles = @(
                if (Test-Path -LiteralPath $suffixFolder -PathType Container) {
                    Get-ChildItem -LiteralPath $suffixFolder -File -ErrorAction Stop |
                        Where-Object { Test-AsiToPixSupportedCalibrationFileName -FileName $_.Name }
                }
            )
            $suffixExposureKeys = @(
                Get-AsiToPixFlatFolderExposure -Path $suffixFolder
            )
            if ($suffixImageFiles.Count -gt 0 -and $exposureKey -notin $suffixExposureKeys) {
                $foundExposureText = if ($suffixExposureKeys.Count -gt 0) {
                    $suffixExposureKeys -join ", "
                } else {
                    "unrecognized"
                }
                throw "Flat collision folder '$suffixFolder' contains exposure(s) $foundExposureText, expected $exposureKey seconds."
            }

            $destinationRecord.DestinationFolder = $suffixFolder
            $splitKey = "$baseFolder|$exposureKey"
            if ($reportedFlatSplits.Add($splitKey)) {
                $flatExposureSplits += [PSCustomObject]@{
                    BaseFolder        = $baseFolder
                    DestinationFolder = $suffixFolder
                    ExposureSeconds   = $exposureKey
                    Suffix            = $exposureSuffix
                }
            }
        }
    }

    $entries = @()
    $plannedPaths = @{}
    foreach ($destinationRecord in $destinationRecords) {
        $record = $destinationRecord.Record
        $destinationFolder = $destinationRecord.DestinationFolder
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
        CalibrationRoot  = $resolvedCalibrationRoot
        SourceCount      = $SourceRecord.Count
        PlannedCount     = @($entries | Where-Object { $_.Status -eq "Planned" }).Count
        ExistingCount    = @($entries | Where-Object { $_.Status -eq "Exists" }).Count
        ConflictCount    = @($entries | Where-Object { $_.Status -eq "Conflict" }).Count
        Entries          = @($entries)
        AdditionWarnings = @($additionWarnings)
        FlatExposureSplits = @($flatExposureSplits)
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
        $SetupName = Get-AsiToPixCalibrationSetupName -SourcePath $resolvedSourcePath
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
    foreach ($split in @($Plan.FlatExposureSplits)) {
        Write-Host "  [flat-set] $($split.ExposureSeconds)s uses neutral folder suffix '$($split.Suffix)':" `
            -ForegroundColor Yellow
        Write-Host "    $($split.DestinationFolder)" -ForegroundColor DarkYellow
    }
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

function Get-AsiToPixCalibrationUncShareKey {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if ($Path -notmatch '^[\\/]{2}(?<server>[^\\/]+)[\\/](?<share>[^\\/]+)') {
        return ""
    }

    return "\\$($Matches['server'].ToLowerInvariant())\$($Matches['share'].ToLowerInvariant())"
}

function Get-AsiToPixCalibrationNetworkLocationKey {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $uncShareKey = Get-AsiToPixCalibrationUncShareKey -Path $Path
    if (-not [string]::IsNullOrWhiteSpace($uncShareKey)) {
        return $uncShareKey
    }

    $pathRoot = [System.IO.Path]::GetPathRoot($Path)
    if ([string]::IsNullOrWhiteSpace($pathRoot) -or $pathRoot -notmatch '^(?<drive>[A-Za-z]):[\\/]$') {
        return ""
    }

    $driveName = $Matches['drive'].ToUpperInvariant()
    $drive = Get-PSDrive -Name $driveName -PSProvider FileSystem -ErrorAction SilentlyContinue
    if ($null -ne $drive) {
        $displayRootProperty = $drive.PSObject.Properties['DisplayRoot']
        if ($null -ne $displayRootProperty -and
            -not [string]::IsNullOrWhiteSpace([string]$displayRootProperty.Value)) {
            $mappedShareKey = Get-AsiToPixCalibrationUncShareKey -Path ([string]$displayRootProperty.Value)
            if (-not [string]::IsNullOrWhiteSpace($mappedShareKey)) {
                return $mappedShareKey
            }
        }
    }

    try {
        $driveInfo = [System.IO.DriveInfo]::new("${driveName}:\")
        if ($driveInfo.DriveType -eq [System.IO.DriveType]::Network) {
            return "drive:$driveName"
        }
    } catch {
        return ""
    }

    return ""
}

function Test-AsiToPixCalibrationUseRobocopy {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$Entry,

        [Parameter(Mandatory = $true)]
        [string]$DestinationRoot
    )

    if ($Entry.Count -eq 0 -or
        $null -eq (Get-Command robocopy.exe -CommandType Application -ErrorAction SilentlyContinue)) {
        return $false
    }

    $destinationKey = Get-AsiToPixCalibrationNetworkLocationKey -Path $DestinationRoot
    if ([string]::IsNullOrWhiteSpace($destinationKey)) {
        return $false
    }

    foreach ($item in $Entry) {
        $sourceKey = Get-AsiToPixCalibrationNetworkLocationKey -Path $item.SourcePath
        if ([string]::IsNullOrWhiteSpace($sourceKey) -or
            -not $sourceKey.Equals($destinationKey, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $false
        }
    }

    return $true
}

function Get-AsiToPixCalibrationRobocopyProgressStatus {
    param(
        [Parameter(Mandatory = $true)]
        [int]$CompletedFileCount,

        [Parameter(Mandatory = $true)]
        [int]$TotalFileCount,

        [Parameter(Mandatory = $true)]
        [long]$CompletedByteCount,

        [Parameter(Mandatory = $true)]
        [timespan]$Elapsed,

        [Parameter(Mandatory = $true)]
        [int]$ThreadCount
    )

    $percentComplete = [int][Math]::Floor(($CompletedFileCount * 100.0) / $TotalFileCount)
    $elapsedSeconds = [Math]::Max($Elapsed.TotalSeconds, 0.001)
    $megabytesPerSecond = ($CompletedByteCount / 1MB) / $elapsedSeconds
    $elapsedText = "{0:00}:{1:00}:{2:00}" -f [Math]::Floor($Elapsed.TotalHours), $Elapsed.Minutes, $Elapsed.Seconds

    return "$CompletedFileCount/$TotalFileCount files ($percentComplete%), " +
        ("{0:0.0} MB/s avg, elapsed {1} (Robocopy /MT:{2})" -f $megabytesPerSecond, $elapsedText, $ThreadCount)
}

function Invoke-AsiToPixCalibrationRobocopyBatch {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourceDirectory,

        [Parameter(Mandatory = $true)]
        [string]$DestinationDirectory,

        [Parameter(Mandatory = $true)]
        [ValidateCount(1, 64)]
        [string[]]$FileName,

        [ValidateRange(1, 16)]
        [int]$ThreadCount = 4,

        [int]$CompletedFileCount = 0,

        [int]$TotalFileCount = 0,

        [long]$CompletedByteCount = 0,

        [long]$BatchByteCount = 0,

        [long]$TotalByteCount = 0,

        [System.Diagnostics.Stopwatch]$CopyStopwatch
    )

    $robocopy = Get-Command robocopy.exe -CommandType Application -ErrorAction Stop
    $arguments = @(
        $SourceDirectory,
        $DestinationDirectory
    ) + $FileName + @(
        "/J",
        "/MT:$ThreadCount",
        "/R:2",
        "/W:1",
        "/COPY:DAT",
        "/DCOPY:DA",
        "/NDL",
        "/NJH",
        "/NJS"
    )

    $robocopyOutput = [System.Collections.Generic.List[string]]::new()
    $observedCompletedFiles = 0
    & $robocopy.Source @arguments 2>&1 | ForEach-Object {
        $outputLine = [string]$_
        [void]$robocopyOutput.Add($outputLine)
        if ($robocopyOutput.Count -gt 20) {
            $robocopyOutput.RemoveAt(0)
        }

        if ($TotalFileCount -gt 0 -and
            $observedCompletedFiles -lt $FileName.Count -and
            $outputLine.Trim() -match '^100(?:[.,]0+)?%$') {
            $observedCompletedFiles++
            $overallCompletedFiles = [Math]::Min(
                $CompletedFileCount + $observedCompletedFiles,
                $TotalFileCount
            )
            $estimatedBatchBytes = [long][Math]::Round(
                $BatchByteCount * ($observedCompletedFiles / [double]$FileName.Count)
            )
            $overallCompletedBytes = [Math]::Min(
                $CompletedByteCount + $estimatedBatchBytes,
                $TotalByteCount
            )
            $percentComplete = [int][Math]::Floor(($overallCompletedFiles * 100.0) / $TotalFileCount)
            $status = Get-AsiToPixCalibrationRobocopyProgressStatus `
                -CompletedFileCount $overallCompletedFiles `
                -TotalFileCount $TotalFileCount `
                -CompletedByteCount $overallCompletedBytes `
                -Elapsed $CopyStopwatch.Elapsed `
                -ThreadCount $ThreadCount
            Write-Progress `
                -Activity "Copying calibration frames from NAS" `
                -Status $status `
                -PercentComplete $percentComplete
        }
    }

    $exitCode = $LASTEXITCODE
    if ($exitCode -ge 8) {
        $details = @($robocopyOutput | Select-Object -Last 10) -join [Environment]::NewLine
        throw "Robocopy failed with exit code $exitCode while copying from '$SourceDirectory' to '$DestinationDirectory'. $details"
    }
}

function Invoke-AsiToPixCalibrationRobocopyImport {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [object[]]$Entry,

        [Parameter(Mandatory = $true)]
        [string]$CalibrationRoot,

        [ValidateRange(1, 16)]
        [int]$ThreadCount = 4
    )

    $stagingName = ".asitopix-calibration-import-$([guid]::NewGuid().ToString('N'))"
    $stagingRoot = Join-Path -Path $CalibrationRoot -ChildPath $stagingName
    New-Item -ItemType Directory -Path $stagingRoot -ErrorAction Stop | Out-Null
    $batchSize = 16
    $totalFiles = $Entry.Count
    $totalBytes = [long](($Entry | Measure-Object -Property SourceLength -Sum).Sum)
    $copiedFiles = 0
    $copiedBytes = [long]0
    $batchNumber = 0
    $copyStopwatch = [System.Diagnostics.Stopwatch]::StartNew()
    $completedEntries = [System.Collections.Generic.List[object]]::new()

    try {
        foreach ($sourceGroup in $Entry | Group-Object { Split-Path -Path $_.SourcePath -Parent }) {
            $groupItems = @($sourceGroup.Group)
            for ($offset = 0; $offset -lt $groupItems.Count; $offset += $batchSize) {
                $batchNumber++
                $lastIndex = [Math]::Min($offset + $batchSize - 1, $groupItems.Count - 1)
                $batchItems = @($groupItems[$offset..$lastIndex])
                $batchBytes = [long](($batchItems | Measure-Object -Property SourceLength -Sum).Sum)
                $batchStagingRoot = Join-Path -Path $stagingRoot -ChildPath ("batch-{0:0000}" -f $batchNumber)
                New-Item -ItemType Directory -Path $batchStagingRoot -ErrorAction Stop | Out-Null

                Invoke-AsiToPixCalibrationRobocopyBatch `
                    -SourceDirectory $sourceGroup.Name `
                    -DestinationDirectory $batchStagingRoot `
                    -FileName @($batchItems | ForEach-Object { Split-Path -Path $_.SourcePath -Leaf }) `
                    -ThreadCount $ThreadCount `
                    -CompletedFileCount $copiedFiles `
                    -TotalFileCount $totalFiles `
                    -CompletedByteCount $copiedBytes `
                    -BatchByteCount $batchBytes `
                    -TotalByteCount $totalBytes `
                    -CopyStopwatch $copyStopwatch

                foreach ($item in $batchItems) {
                    $fileName = Split-Path -Path $item.SourcePath -Leaf
                    $stagedFile = Join-Path -Path $batchStagingRoot -ChildPath $fileName
                    if (-not (Test-Path -LiteralPath $stagedFile -PathType Leaf)) {
                        throw "Robocopy did not create the expected staged file: '$stagedFile'."
                    }

                    $stagedFileInfo = Get-Item -LiteralPath $stagedFile -ErrorAction Stop
                    if ($stagedFileInfo.Length -ne $item.SourceLength) {
                        throw "Staged file size mismatch for '$stagedFile': expected $($item.SourceLength) byte(s), got $($stagedFileInfo.Length)."
                    }
                    if (Test-Path -LiteralPath $item.DestinationPath) {
                        throw "Destination path appeared during import and will not be overwritten: $($item.DestinationPath)"
                    }

                    Move-Item -LiteralPath $stagedFile -Destination $item.DestinationPath -ErrorAction Stop
                    $completedEntries.Add($item)
                }

                Remove-Item -LiteralPath $batchStagingRoot -Force -ErrorAction Stop
                $copiedFiles += $batchItems.Count
                $copiedBytes += $batchBytes
                $percentComplete = [int][Math]::Floor(($copiedFiles * 100.0) / $totalFiles)
                $status = Get-AsiToPixCalibrationRobocopyProgressStatus `
                    -CompletedFileCount $copiedFiles `
                    -TotalFileCount $totalFiles `
                    -CompletedByteCount $copiedBytes `
                    -Elapsed $copyStopwatch.Elapsed `
                    -ThreadCount $ThreadCount
                Write-Progress `
                    -Activity "Copying calibration frames from NAS" `
                    -Status $status `
                    -PercentComplete $percentComplete
            }
        }

        return @($completedEntries)
    } catch {
        throw "Fast NAS calibration copy failed. Recoverable staged files, if any, remain under '$stagingRoot'. $($_.Exception.Message)"
    } finally {
        $copyStopwatch.Stop()
        Write-Progress -Activity "Copying calibration frames from NAS" -Completed
        if (Test-Path -LiteralPath $stagingRoot -PathType Container) {
            $remainingItems = @(Get-ChildItem -LiteralPath $stagingRoot -Force -ErrorAction Stop)
            if ($remainingItems.Count -eq 0) {
                Remove-Item -LiteralPath $stagingRoot -Force -ErrorAction Stop
            }
        }
    }
}

function Invoke-AsiToPixCalibrationImportPlan {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Plan
    )

    $approvedEntries = [System.Collections.Generic.List[object]]::new()
    $whatIfCount = 0
    foreach ($entry in @($Plan.Entries | Where-Object { $_.Status -eq "Planned" })) {
        if (-not $PSCmdlet.ShouldProcess(
            $entry.DestinationPath,
            "Copy calibration file from '$($entry.SourcePath)'"
        )) {
            $whatIfCount++
            continue
        }

        $approvedEntries.Add($entry)
    }

    $preparedDirectories = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    foreach ($entry in $approvedEntries) {
        if ($preparedDirectories.Add($entry.DestinationFolder)) {
            if (Test-Path -LiteralPath $entry.DestinationFolder) {
                if (-not (Test-Path -LiteralPath $entry.DestinationFolder -PathType Container)) {
                    throw "Calibration destination folder path is occupied by a non-directory item: $($entry.DestinationFolder)"
                }
            } else {
                New-Item -ItemType Directory -Path $entry.DestinationFolder -Force -ErrorAction Stop | Out-Null
            }
        }

        if (Test-Path -LiteralPath $entry.DestinationPath) {
            throw "Calibration destination file appeared after planning and will not be overwritten: $($entry.DestinationPath)"
        }
    }

    $copyEngine = "CopyItem"
    $completedEntries = @()
    if ($approvedEntries.Count -gt 0) {
        if (Test-AsiToPixCalibrationUseRobocopy `
            -Entry @($approvedEntries) `
            -DestinationRoot $Plan.CalibrationRoot) {
            $copyEngine = "Robocopy"
            Write-Host "[INFO] Same-share NAS calibration copy detected; using Robocopy /J /MT:4 with safe staging." `
                -ForegroundColor Cyan
            $completedEntries = @(
                Invoke-AsiToPixCalibrationRobocopyImport `
                    -Entry @($approvedEntries) `
                    -CalibrationRoot $Plan.CalibrationRoot `
                    -ThreadCount 4
            )
        } else {
            $completedEntries = @(
                foreach ($entry in $approvedEntries) {
                    try {
                        if (Test-Path -LiteralPath $entry.DestinationPath) {
                            throw "Destination file appeared after planning; refusing to overwrite it."
                        }
                        Copy-Item `
                            -LiteralPath $entry.SourcePath `
                            -Destination $entry.DestinationPath `
                            -ErrorAction Stop
                        $entry
                    } catch {
                        throw "Failed to copy calibration file '$($entry.SourcePath)' to '$($entry.DestinationPath)': $($_.Exception.Message)"
                    }
                }
            )
        }
    }

    foreach ($entry in $completedEntries) {
        Write-Host "  [+] $(Split-Path -Path $entry.SourcePath -Leaf) -> $($entry.DestinationFolder)" `
            -ForegroundColor Green
    }

    return [PSCustomObject]@{
        CopiedCount   = $completedEntries.Count
        ExistingCount = $Plan.ExistingCount
        ConflictCount = $Plan.ConflictCount
        WhatIfCount   = $whatIfCount
        CopyEngine    = $copyEngine
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

        [string]$AngleDegrees = "",

        [switch]$SkipConfirmation
    )

    $resolvedSourcePath = (Resolve-Path -LiteralPath $SourcePath).ProviderPath
    $resolvedCalibrationRoot = (Resolve-Path -LiteralPath $CalibrationRoot).ProviderPath
    $setupName = Get-AsiToPixCalibrationSetupName -SourcePath $resolvedSourcePath
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
        -not $SkipConfirmation -and
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
    Find-AsiToPixCalibrationImportFolder, `
    Get-AsiToPixCalibrationImportPlan, `
    Import-AsiToPixCalibration, `
    Invoke-AsiToPixCalibrationImportPlan, `
    Read-AsiToPixCalibrationConfirmation, `
    Test-AsiToPixSupportedCalibrationFileName
