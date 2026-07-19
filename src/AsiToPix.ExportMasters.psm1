Set-StrictMode -Version Latest

$environmentModule = Join-Path -Path $PSScriptRoot -ChildPath "AsiToPix.Environment.psm1"
Import-Module $environmentModule

$projectMetadataModule = Join-Path -Path $PSScriptRoot -ChildPath "AsiToPix.ProjectMetadata.psm1"
Import-Module $projectMetadataModule

$namesModule = Join-Path -Path $PSScriptRoot -ChildPath "AsiToPix.Names.psm1"
Import-Module $namesModule

$frameFoldersModule = Join-Path -Path $PSScriptRoot -ChildPath "AsiToPix.FrameFolders.psm1"
Import-Module $frameFoldersModule -Force

function Get-AsiToPixObjectPropertyValue {
    param(
        [AllowNull()]
        [object]$InputObject,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if ($null -eq $InputObject) {
        return $null
    }

    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $null
    }

    return $property.Value
}

function ConvertTo-AsiToPixNumericKey {
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Value,

        [AllowEmptyString()]
        [string]$UnitPattern = ""
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $null
    }

    $numericText = $Value.Trim()
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
        return $null
    }

    $canonicalText = $number.ToString("G29", [System.Globalization.CultureInfo]::InvariantCulture)
    return (Get-AsiToPixCanonicalNumericText -Value $canonicalText)
}

function Get-AsiToPixRegexValue {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Text,

        [Parameter(Mandatory = $true)]
        [string]$Pattern,

        [Parameter(Mandatory = $true)]
        [string]$GroupName,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[string]]$Issue,

        [Parameter(Mandatory = $true)]
        [string]$Description
    )

    $values = @(
        [regex]::Matches($Text, $Pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase) |
            ForEach-Object { $_.Groups[$GroupName].Value } |
            Where-Object { -not [string]::IsNullOrWhiteSpace($_) } |
            Select-Object -Unique
    )

    if ($values.Count -gt 1) {
        $Issue.Add("Conflicting $Description values: $($values -join ', ').")
        return $null
    }

    if ($values.Count -eq 1) {
        return [string]$values[0]
    }

    return $null
}

function Get-AsiToPixCleanMasterFileName {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("Bias", "Dark", "Flat")]
        [string]$MasterType,

        [AllowNull()]
        [AllowEmptyString()]
        [string]$Bin,

        [AllowNull()]
        [AllowEmptyString()]
        [string]$Resolution
    )

    $parts = [System.Collections.Generic.List[string]]::new()
    $parts.Add("master$MasterType")
    if (-not [string]::IsNullOrWhiteSpace($Bin)) {
        $parts.Add("BIN-$Bin")
    }
    if (-not [string]::IsNullOrWhiteSpace($Resolution)) {
        $parts.Add($Resolution)
    }

    return "$(($parts.ToArray()) -join '_').xisf"
}

function Get-AsiToPixWbppMasterInfo {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileName
    )

    $stem = [System.IO.Path]::GetFileNameWithoutExtension($FileName)
    $masterType = if ($stem -match '^(?i:masterBias)(?:_|$)') {
        "Bias"
    } elseif ($stem -match '^(?i:masterDark)(?:_|$)') {
        "Dark"
    } elseif ($stem -match '^(?i:masterFlat)(?:_|$)') {
        "Flat"
    } else {
        return $null
    }

    $issues = [System.Collections.Generic.List[string]]::new()
    $gainText = Get-AsiToPixRegexValue `
        -Text $stem `
        -Pattern '(?:^|_)GAIN-(?<value>\d+(?:\.\d+)?)(?=_|$)' `
        -GroupName "value" `
        -Issue $issues `
        -Description "GAIN"
    $temperatureText = Get-AsiToPixRegexValue `
        -Text $stem `
        -Pattern '(?:^|_)TEMP-(?<value>-?\d+(?:\.\d+)?C)(?=_|$)' `
        -GroupName "value" `
        -Issue $issues `
        -Description "TEMP"
    $filter = Get-AsiToPixRegexValue `
        -Text $stem `
        -Pattern '(?:^|_)FILTER-(?<value>[^_]+)(?=_|$)' `
        -GroupName "value" `
        -Issue $issues `
        -Description "FILTER"
    $camera = Get-AsiToPixRegexValue `
        -Text $stem `
        -Pattern '(?:^|_)CAM-(?<value>[^_]+)(?=_|$)' `
        -GroupName "value" `
        -Issue $issues `
        -Description "CAM"
    $bin = Get-AsiToPixRegexValue `
        -Text $stem `
        -Pattern '(?:^|_)BIN-(?<value>\d+)(?=_|$)' `
        -GroupName "value" `
        -Issue $issues `
        -Description "BIN"
    $resolution = Get-AsiToPixRegexValue `
        -Text $stem `
        -Pattern '(?:^|_)(?<value>\d+x\d+)(?=_|$)' `
        -GroupName "value" `
        -Issue $issues `
        -Description "resolution"

    $exposureValues = [System.Collections.Generic.List[string]]::new()
    foreach ($pattern in @(
        '(?:^|_)EXP-(?<value>\d+(?:\.\d+)?s)(?=_|$)',
        '(?:^|_)EXPOSURE-(?<value>\d+(?:\.\d+)?s)(?=_|$)'
    )) {
        foreach ($match in [regex]::Matches($stem, $pattern, [System.Text.RegularExpressions.RegexOptions]::IgnoreCase)) {
            $numericExposure = ConvertTo-AsiToPixNumericKey -Value $match.Groups["value"].Value -UnitPattern 's'
            if ($null -ne $numericExposure -and -not $exposureValues.Contains($numericExposure)) {
                $exposureValues.Add($numericExposure)
            }
        }
    }

    if ($exposureValues.Count -gt 1) {
        $issues.Add("Conflicting EXP/EXPOSURE values: $($exposureValues -join ', ').")
    }
    $exposure = if ($exposureValues.Count -eq 1) { $exposureValues[0] } else { $null }

    $gain = ConvertTo-AsiToPixNumericKey -Value $gainText
    $temperature = ConvertTo-AsiToPixNumericKey -Value $temperatureText -UnitPattern 'C'
    if ($null -ne $filter) {
        $filter = $filter.ToUpperInvariant()
    }

    return [PSCustomObject]@{
        MasterType    = $masterType
        Gain          = $gain
        Temperature   = $temperature
        Exposure      = $exposure
        Filter        = $filter
        Camera        = $camera
        Bin           = $bin
        Resolution    = $resolution
        CleanFileName = Get-AsiToPixCleanMasterFileName -MasterType $masterType -Bin $bin -Resolution $resolution
        Issues        = @($issues)
    }
}

function Test-AsiToPixProjectMetadataNeedsAstroPhotoRoot {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Metadata
    )

    foreach ($camera in @(Get-AsiToPixObjectPropertyValue -InputObject $Metadata -Name "Cameras")) {
        $folders = Get-AsiToPixObjectPropertyValue -InputObject $camera -Name "CalibrationFolders"
        foreach ($propertyName in @("Darks", "Biases", "Flats", "FlatDarks")) {
            $path = Get-AsiToPixObjectPropertyValue -InputObject $folders -Name $propertyName
            if ([string]::IsNullOrWhiteSpace([string]$path)) {
                return $true
            }
        }
    }

    return $false
}

function Assert-AsiToPixProjectMetadata {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Metadata,

        [Parameter(Mandatory = $true)]
        [string]$MetadataPath
    )

    $pixPath = [string](Get-AsiToPixObjectPropertyValue -InputObject $Metadata -Name "PixPath")
    if ([string]::IsNullOrWhiteSpace($pixPath)) {
        throw "Project metadata '$MetadataPath' is missing the required PixPath value."
    }

    $scope = [string](Get-AsiToPixObjectPropertyValue -InputObject $Metadata -Name "Scope")
    if ([string]::IsNullOrWhiteSpace($scope)) {
        throw "Project metadata '$MetadataPath' is missing the required Scope value."
    }

    $cameras = @(Get-AsiToPixObjectPropertyValue -InputObject $Metadata -Name "Cameras")
    if ($cameras.Count -eq 0) {
        throw "Project metadata '$MetadataPath' must contain at least one Cameras entry."
    }

    foreach ($camera in $cameras) {
        $cameraName = [string](Get-AsiToPixObjectPropertyValue -InputObject $camera -Name "Name")
        if ([string]::IsNullOrWhiteSpace($cameraName)) {
            throw "Project metadata '$MetadataPath' contains a Cameras entry without the required Name value."
        }
    }
}

function Read-AsiToPixProjectMetadata {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    if ([string]::IsNullOrWhiteSpace($Path)) {
        throw "The project metadata path cannot be empty."
    }
    if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
        throw "Project metadata file not found: '$Path'."
    }

    $resolvedPath = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).ProviderPath
    try {
        $metadata = Get-Content -LiteralPath $resolvedPath -Raw -Encoding UTF8 -ErrorAction Stop |
            ConvertFrom-Json -ErrorAction Stop
    } catch {
        throw "Could not read project metadata '$resolvedPath': $($_.Exception.Message)"
    }

    Assert-AsiToPixProjectMetadata -Metadata $metadata -MetadataPath $resolvedPath
    return $metadata
}

function Find-AsiToPixProcessingProjectMetadata {
    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProcessingRoot,

        [Parameter(Mandatory = $true)]
        [string]$ObjectName
    )

    if (-not (Test-Path -LiteralPath $ProcessingRoot -PathType Container)) {
        throw "AstroPhoto Processing folder not found: '$ProcessingRoot'."
    }

    $resolvedProcessingRoot = (Resolve-Path -LiteralPath $ProcessingRoot -ErrorAction Stop).ProviderPath
    $objectFolders = @(Get-ChildItem -LiteralPath $resolvedProcessingRoot -Directory -ErrorAction Stop)
    $objectMatches = @(Get-AsiToPixNameMatch `
        -DetectedName $ObjectName `
        -Candidates @($objectFolders | Select-Object -ExpandProperty Name))
    if ($objectMatches.Count -eq 0) {
        return @()
    }

    $projects = foreach ($objectMatch in $objectMatches) {
        $objectFolder = $objectFolders |
            Where-Object { $_.Name -eq $objectMatch.Name } |
            Select-Object -First 1
        if ($null -eq $objectFolder) {
            continue
        }

        foreach ($metadataFile in @(Get-ChildItem `
            -LiteralPath $objectFolder.FullName `
            -File `
            -Filter "project_meta.json" `
            -Recurse `
            -ErrorAction Stop)) {
            [PSCustomObject]@{
                ObjectName  = $objectFolder.Name
                ProjectName = $metadataFile.Directory.Name
                ProjectPath = $metadataFile.Directory.FullName
                MetaPath    = $metadataFile.FullName
                Score       = $objectMatch.Score
            }
        }
    }

    return @($projects | Sort-Object `
        -Property @{ Expression = "Score"; Descending = $true }, ObjectName, ProjectName, MetaPath)
}

function Resolve-AsiToPixProjectMetadataPath {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$InputValue,

        [AllowEmptyString()]
        [string]$ProcessingRoot = "",

        [scriptblock]$SelectionReader = {
            param($Prompt)
            Read-Host $Prompt
        }
    )

    $resolvedInput = $InputValue.Trim().Trim('"')
    if ([string]::IsNullOrWhiteSpace($resolvedInput)) {
        throw "The project metadata path or object name cannot be empty."
    }

    if (Test-Path -LiteralPath $resolvedInput -PathType Leaf) {
        return (Resolve-Path -LiteralPath $resolvedInput -ErrorAction Stop).ProviderPath
    }

    if (Test-Path -LiteralPath $resolvedInput -PathType Container) {
        $metadataPath = Join-Path -Path $resolvedInput -ChildPath "project_meta.json"
        if (-not (Test-Path -LiteralPath $metadataPath -PathType Leaf)) {
            throw "Project folder '$resolvedInput' does not contain project_meta.json."
        }

        return (Resolve-Path -LiteralPath $metadataPath -ErrorAction Stop).ProviderPath
    }

    $looksLikePath = [System.IO.Path]::IsPathRooted($resolvedInput) -or
        $resolvedInput.Contains([System.IO.Path]::DirectorySeparatorChar) -or
        $resolvedInput.Contains([System.IO.Path]::AltDirectorySeparatorChar) -or
        [System.IO.Path]::GetExtension($resolvedInput) -ieq ".json"
    if ($looksLikePath) {
        throw "Project metadata path not found: '$resolvedInput'."
    }

    if ([string]::IsNullOrWhiteSpace($ProcessingRoot)) {
        throw "ProcessingRoot is required to resolve object name '$resolvedInput'."
    }

    $candidates = @(Find-AsiToPixProcessingProjectMetadata `
        -ProcessingRoot $ProcessingRoot `
        -ObjectName $resolvedInput)
    if ($candidates.Count -eq 0) {
        throw "No processing project matching object name '$resolvedInput' with project_meta.json was found under '$ProcessingRoot'."
    }

    if ($candidates.Count -eq 1) {
        Write-Host "[INFO] Object '$resolvedInput' resolved to project: $($candidates[0].ProjectPath)" -ForegroundColor Cyan
        return $candidates[0].MetaPath
    }

    Write-Host "`nMatching processing projects for '${resolvedInput}':" -ForegroundColor Cyan
    for ($i = 0; $i -lt $candidates.Count; $i++) {
        $candidate = $candidates[$i]
        Write-Host (" [{0}] {1} | {2}" -f ($i + 1), $candidate.ObjectName, $candidate.ProjectName) -ForegroundColor White
        Write-Host "     $($candidate.MetaPath)" -ForegroundColor DarkGray
    }

    do {
        $answerValue = & $SelectionReader "Select processing project index, press Enter for 1, or type a project_meta.json path"
        if ($null -eq $answerValue) {
            throw "Project selection input ended before an answer was received."
        }

        $answer = ([string]$answerValue).Trim().Trim('"')
        if ([string]::IsNullOrWhiteSpace($answer)) {
            return $candidates[0].MetaPath
        }

        if ($answer -match '^\d+$') {
            $index = [int]$answer - 1
            if ($index -ge 0 -and $index -lt $candidates.Count) {
                return $candidates[$index].MetaPath
            }
        } elseif (Test-Path -LiteralPath $answer -PathType Leaf) {
            return (Resolve-Path -LiteralPath $answer -ErrorAction Stop).ProviderPath
        } elseif (Test-Path -LiteralPath $answer -PathType Container) {
            $selectedMetadataPath = Join-Path -Path $answer -ChildPath "project_meta.json"
            if (Test-Path -LiteralPath $selectedMetadataPath -PathType Leaf) {
                return (Resolve-Path -LiteralPath $selectedMetadataPath -ErrorAction Stop).ProviderPath
            }
        }

        Write-Host "[!] Invalid processing project selection." -ForegroundColor Red
    } while ($true)
}

function Get-AsiToPixCalibrationFolder {
    param(
        [Parameter(Mandatory = $true)]
        [object]$CameraMetadata,

        [Parameter(Mandatory = $true)]
        [ValidateSet("Biases", "Darks", "Flats", "FlatDarks")]
        [string]$Category,

        [AllowNull()]
        [AllowEmptyString()]
        [string]$AstroPhotoRoot
    )

    $folders = Get-AsiToPixObjectPropertyValue -InputObject $CameraMetadata -Name "CalibrationFolders"
    $configuredPath = [string](Get-AsiToPixObjectPropertyValue -InputObject $folders -Name $Category)
    if (-not [string]::IsNullOrWhiteSpace($configuredPath)) {
        return [System.IO.Path]::GetFullPath($configuredPath)
    }

    if ([string]::IsNullOrWhiteSpace($AstroPhotoRoot)) {
        $cameraName = [string](Get-AsiToPixObjectPropertyValue -InputObject $CameraMetadata -Name "Name")
        throw "CalibrationFolders.$Category is missing for camera '$cameraName', and no AstroPhoto root was provided."
    }

    $folderKind = switch ($Category) {
        "Biases" { "Bias" }
        "Darks" { "Dark" }
        "Flats" { "Flat" }
        "FlatDarks" { "FlatDark" }
    }
    $cameraName = [string](Get-AsiToPixObjectPropertyValue -InputObject $CameraMetadata -Name "Name")
    $calibrationRoot = Join-Path -Path $AstroPhotoRoot -ChildPath "Calibration"
    $cameraRoot = Join-Path -Path $calibrationRoot -ChildPath $cameraName
    $masterRoot = Join-Path -Path $cameraRoot -ChildPath "Master"
    $existingFolders = @(Get-AsiToPixChildFrameFolder -Path $masterRoot -Kind $folderKind)
    if ($existingFolders.Count -gt 1) {
        throw "Multiple $Category folders were found under '$masterRoot': $($existingFolders.FullName -join ', ')."
    }
    if ($existingFolders.Count -eq 1) {
        return $existingFolders[0].FullName
    }

    $categoryFolder = Get-AsiToPixCanonicalFrameFolderName -Kind $folderKind
    return (Join-Path -Path $masterRoot -ChildPath $categoryFolder)
}

function Get-AsiToPixSourceRootFromMasterRoot {
    param(
        [Parameter(Mandatory = $true)]
        [string]$MasterRoot
    )

    $normalizedMasterRoot = [System.IO.Path]::GetFullPath($MasterRoot)
    $kindName = Split-Path -Path $normalizedMasterRoot -Leaf
    $masterContainer = Split-Path -Path $normalizedMasterRoot -Parent
    if ((Split-Path -Path $masterContainer -Leaf) -ine "Master") {
        throw "Calibration master path must have the form '<camera>\Master\<kind>': '$MasterRoot'."
    }

    $cameraRoot = Split-Path -Path $masterContainer -Parent
    $sourceContainer = Join-Path -Path $cameraRoot -ChildPath "Source"
    return (Join-Path -Path $sourceContainer -ChildPath $kindName)
}

function Get-AsiToPixRelativeChildPath {
    param(
        [Parameter(Mandatory = $true)]
        [string]$RootPath,

        [Parameter(Mandatory = $true)]
        [string]$ChildPath
    )

    $separator = [System.IO.Path]::DirectorySeparatorChar
    $alternateSeparator = [System.IO.Path]::AltDirectorySeparatorChar
    $root = [System.IO.Path]::GetFullPath($RootPath).Replace($alternateSeparator, $separator).TrimEnd($separator)
    $child = [System.IO.Path]::GetFullPath($ChildPath).Replace($alternateSeparator, $separator).TrimEnd($separator)

    if ($child.Equals($root, [System.StringComparison]::OrdinalIgnoreCase)) {
        return ""
    }

    $prefix = "$root$separator"
    if (-not $child.StartsWith($prefix, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $null
    }

    return $child.Substring($prefix.Length)
}

function Get-AsiToPixSourceTagInfo {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [ValidateSet("Biases", "Darks", "Flats", "FlatDarks")]
        [string]$Category
    )

    $issues = [System.Collections.Generic.List[string]]::new()
    $gainText = Get-AsiToPixRegexValue -Text $Name -Pattern '(?:^|_)Gain_(?<value>[^_]+)(?=_|$)' -GroupName "value" -Issue $issues -Description "Gain"
    $temperatureText = Get-AsiToPixRegexValue -Text $Name -Pattern '(?:^|_)Temp_(?<value>-?\d+(?:\.\d+)?C)(?=_|$)' -GroupName "value" -Issue $issues -Description "Temp"
    $exposureText = Get-AsiToPixRegexValue -Text $Name -Pattern '(?:^|_)Exp_(?<value>\d+(?:\.\d+)?s)(?=_|$)' -GroupName "value" -Issue $issues -Description "Exp"
    $filter = Get-AsiToPixRegexValue -Text $Name -Pattern '(?:^|_)Filter_(?<value>[^_]+)(?=_|$)' -GroupName "value" -Issue $issues -Description "Filter"
    $camera = Get-AsiToPixRegexValue -Text $Name -Pattern '(?:^|_)Cam_(?<value>.+)$' -GroupName "value" -Issue $issues -Description "Cam"

    if ($null -ne $filter) {
        $filter = $filter.ToUpperInvariant()
    }

    return [PSCustomObject]@{
        Category    = $Category
        Gain        = ConvertTo-AsiToPixNumericKey -Value $gainText
        Temperature = ConvertTo-AsiToPixNumericKey -Value $temperatureText -UnitPattern 'C'
        Exposure    = ConvertTo-AsiToPixNumericKey -Value $exposureText -UnitPattern 's'
        Filter      = $filter
        Camera      = $camera
        Issues      = @($issues)
    }
}

function Get-AsiToPixLinkTargetPath {
    param(
        [Parameter(Mandatory = $true)]
        [System.IO.DirectoryInfo]$Link
    )

    if ($Link.LinkType -ne "SymbolicLink") {
        throw "Project source entry is not a symbolic link: '$($Link.FullName)'."
    }

    $targets = @($Link.Target | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
    if ($targets.Count -ne 1) {
        throw "Project source link must have exactly one target: '$($Link.FullName)'."
    }

    $targetPath = [string]$targets[0]
    if (-not [System.IO.Path]::IsPathRooted($targetPath)) {
        $targetPath = Join-Path -Path $Link.Parent.FullName -ChildPath $targetPath
    }
    if (-not (Test-Path -LiteralPath $targetPath -PathType Container)) {
        throw "Project source link target folder not found: '$targetPath' (link '$($Link.FullName)')."
    }

    return (Resolve-Path -LiteralPath $targetPath -ErrorAction Stop).ProviderPath
}

function Get-AsiToPixProjectSourceRecord {
    param(
        [Parameter(Mandatory = $true)]
        [string]$ProjectSourcePath,

        [Parameter(Mandatory = $true)]
        [object[]]$CameraMetadata,

        [AllowNull()]
        [AllowEmptyString()]
        [string]$AstroPhotoRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[string]]$Diagnostic
    )

    $records = [System.Collections.Generic.List[object]]::new()
    foreach ($category in @("Biases", "Darks", "Flats", "FlatDarks")) {
        $folderKind = switch ($category) {
            "Biases" { "Bias" }
            "Darks" { "Dark" }
            "Flats" { "Flat" }
            "FlatDarks" { "FlatDark" }
        }
        $categoryFolders = @(Get-AsiToPixChildFrameFolder -Path $ProjectSourcePath -Kind $folderKind)

        foreach ($categoryFolder in $categoryFolders) {
            $categoryPath = $categoryFolder.FullName
            foreach ($link in @(Get-ChildItem -LiteralPath $categoryPath -Directory -Force -ErrorAction Stop)) {
                try {
                    $tag = Get-AsiToPixSourceTagInfo -Name $link.Name -Category $category
                    if ($tag.Issues.Count -gt 0) {
                        throw "Could not parse project source link '$($link.FullName)': $($tag.Issues -join ' ')"
                    }
                    if ([string]::IsNullOrWhiteSpace([string]$tag.Camera)) {
                        throw "Project source link has no Cam tag: '$($link.FullName)'."
                    }

                    $camera = @($CameraMetadata | Where-Object {
                        $name = [string](Get-AsiToPixObjectPropertyValue -InputObject $_ -Name "Name")
                        $name.Equals([string]$tag.Camera, [System.StringComparison]::OrdinalIgnoreCase)
                    })
                    if ($camera.Count -ne 1) {
                        throw "Project source link camera '$($tag.Camera)' does not identify exactly one camera in project metadata: '$($link.FullName)'."
                    }

                    if ($category -in @("Darks", "FlatDarks") -and
                        ($null -eq $tag.Gain -or $null -eq $tag.Temperature -or $null -eq $tag.Exposure)) {
                        throw "Project source link is missing Gain, Temp, or Exp metadata: '$($link.FullName)'."
                    }
                    if ($category -eq "Biases" -and ($null -eq $tag.Gain -or $null -eq $tag.Temperature)) {
                        throw "Project source link is missing Gain or Temp metadata: '$($link.FullName)'."
                    }
                    if ($category -eq "Flats" -and [string]::IsNullOrWhiteSpace([string]$tag.Filter)) {
                        throw "Project source link is missing Filter metadata: '$($link.FullName)'."
                    }

                    $targetPath = Get-AsiToPixLinkTargetPath -Link $link
                    $masterRoot = Get-AsiToPixCalibrationFolder -CameraMetadata $camera[0] -Category $category -AstroPhotoRoot $AstroPhotoRoot
                    $sourceRoot = Get-AsiToPixSourceRootFromMasterRoot -MasterRoot $masterRoot
                    $relativeTargetPath = Get-AsiToPixRelativeChildPath -RootPath $sourceRoot -ChildPath $targetPath
                    if ($null -eq $relativeTargetPath) {
                        throw "Project source link target '$targetPath' is outside the expected source root '$sourceRoot'."
                    }
                    $destinationFolder = if ([string]::IsNullOrWhiteSpace($relativeTargetPath)) {
                        $masterRoot
                    } else {
                        Join-Path -Path $masterRoot -ChildPath $relativeTargetPath
                    }

                    $records.Add([PSCustomObject]@{
                        Category          = $category
                        Camera            = [string]$tag.Camera
                        Gain              = $tag.Gain
                        Temperature       = $tag.Temperature
                        Exposure          = $tag.Exposure
                        Filter            = $tag.Filter
                        LinkPath          = $link.FullName
                        TargetPath        = $targetPath
                        DestinationFolder = $destinationFolder
                    })
                } catch {
                    $Diagnostic.Add($_.Exception.Message)
                }
            }
        }
    }

    return @($records)
}

function Get-AsiToPixMetadataCalibrationSourceRecord {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$MetadataSource,

        [Parameter(Mandatory = $true)]
        [object[]]$CameraMetadata,

        [AllowNull()]
        [AllowEmptyString()]
        [string]$AstroPhotoRoot,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [System.Collections.Generic.List[string]]$Diagnostic
    )

    $records = [System.Collections.Generic.List[object]]::new()
    foreach ($metadataRecord in $MetadataSource) {
        try {
            $category = [string](Get-AsiToPixObjectPropertyValue -InputObject $metadataRecord -Name "Type")
            $cameraName = [string](Get-AsiToPixObjectPropertyValue -InputObject $metadataRecord -Name "Camera")
            $sourcePath = [string](Get-AsiToPixObjectPropertyValue -InputObject $metadataRecord -Name "SourcePath")
            $destinationFolder = [string](Get-AsiToPixObjectPropertyValue -InputObject $metadataRecord -Name "DestinationFolder")
            $tag = [string](Get-AsiToPixObjectPropertyValue -InputObject $metadataRecord -Name "Tag")

            if ($category -notin @("Biases", "Darks", "Flats", "FlatDarks")) {
                throw "CalibrationSources entry has an unsupported Type '$category'."
            }
            if ([string]::IsNullOrWhiteSpace($cameraName)) {
                throw "CalibrationSources entry '$tag' is missing Camera."
            }
            if ([string]::IsNullOrWhiteSpace($sourcePath)) {
                throw "CalibrationSources entry '$tag' for camera '$cameraName' is missing SourcePath."
            }
            if ([string]::IsNullOrWhiteSpace($destinationFolder)) {
                throw "CalibrationSources entry '$tag' for camera '$cameraName' is missing DestinationFolder."
            }
            if (-not [System.IO.Path]::IsPathRooted($sourcePath)) {
                throw "CalibrationSources SourcePath must be absolute: '$sourcePath'."
            }
            if (-not [System.IO.Path]::IsPathRooted($destinationFolder)) {
                throw "CalibrationSources DestinationFolder must be absolute: '$destinationFolder'."
            }

            $camera = @($CameraMetadata | Where-Object {
                $name = [string](Get-AsiToPixObjectPropertyValue -InputObject $_ -Name "Name")
                $name.Equals($cameraName, [System.StringComparison]::OrdinalIgnoreCase)
            })
            if ($camera.Count -ne 1) {
                throw "CalibrationSources entry '$tag' identifies camera '$cameraName', which does not match exactly one Cameras entry."
            }

            $masterRoot = Get-AsiToPixCalibrationFolder `
                -CameraMetadata $camera[0] `
                -Category $category `
                -AstroPhotoRoot $AstroPhotoRoot
            $normalizedDestination = [System.IO.Path]::GetFullPath($destinationFolder)
            if ($null -eq (Get-AsiToPixRelativeChildPath -RootPath $masterRoot -ChildPath $normalizedDestination)) {
                throw "CalibrationSources destination '$destinationFolder' is outside the configured master root '$masterRoot'."
            }

            $gain = ConvertTo-AsiToPixNumericKey `
                -Value ([string](Get-AsiToPixObjectPropertyValue -InputObject $metadataRecord -Name "Gain"))
            $temperature = ConvertTo-AsiToPixNumericKey `
                -Value ([string](Get-AsiToPixObjectPropertyValue -InputObject $metadataRecord -Name "TemperatureC"))
            $exposure = ConvertTo-AsiToPixNumericKey `
                -Value ([string](Get-AsiToPixObjectPropertyValue -InputObject $metadataRecord -Name "ExposureSeconds"))
            $filter = [string](Get-AsiToPixObjectPropertyValue -InputObject $metadataRecord -Name "Filter")
            if (-not [string]::IsNullOrWhiteSpace($filter)) {
                $filter = $filter.ToUpperInvariant()
            }

            if ($category -in @("Darks", "FlatDarks") -and
                ($null -eq $gain -or $null -eq $temperature -or $null -eq $exposure)) {
                throw "CalibrationSources entry '$tag' is missing Gain, TemperatureC, or ExposureSeconds."
            }
            if ($category -eq "Biases" -and ($null -eq $gain -or $null -eq $temperature)) {
                throw "CalibrationSources entry '$tag' is missing Gain or TemperatureC."
            }
            if ($category -eq "Flats" -and
                ($null -eq $gain -or $null -eq $temperature -or [string]::IsNullOrWhiteSpace($filter))) {
                throw "CalibrationSources entry '$tag' is missing Gain, TemperatureC, or Filter."
            }

            $records.Add([PSCustomObject]@{
                Category          = $category
                Camera            = $cameraName
                Gain              = $gain
                Temperature       = $temperature
                Exposure          = $exposure
                Filter            = $filter
                LinkPath          = $null
                TargetPath        = [System.IO.Path]::GetFullPath($sourcePath)
                DestinationFolder = $normalizedDestination
            })
        } catch {
            $Diagnostic.Add($_.Exception.Message)
        }
    }

    return @($records)
}

function Test-AsiToPixMasterSourceMatch {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Master,

        [Parameter(Mandatory = $true)]
        [object]$SourceRecord
    )

    if (-not [string]::IsNullOrWhiteSpace([string]$Master.Camera) -and
        -not ([string]$Master.Camera).Equals([string]$SourceRecord.Camera, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $false
    }

    switch ($Master.MasterType) {
        "Bias" {
            return $SourceRecord.Category -eq "Biases" -and
                $null -ne $Master.Gain -and $Master.Gain -eq $SourceRecord.Gain -and
                $null -ne $Master.Temperature -and $Master.Temperature -eq $SourceRecord.Temperature
        }
        "Dark" {
            return $SourceRecord.Category -in @("Darks", "FlatDarks") -and
                $null -ne $Master.Gain -and $Master.Gain -eq $SourceRecord.Gain -and
                $null -ne $Master.Temperature -and $Master.Temperature -eq $SourceRecord.Temperature -and
                $null -ne $Master.Exposure -and $Master.Exposure -eq $SourceRecord.Exposure
        }
        "Flat" {
            if ($SourceRecord.Category -ne "Flats" -or
                [string]::IsNullOrWhiteSpace([string]$Master.Filter) -or
                $Master.Filter -ne $SourceRecord.Filter) {
                return $false
            }
            if ($null -ne $Master.Gain -and $Master.Gain -ne $SourceRecord.Gain) {
                return $false
            }
            if ($null -ne $Master.Temperature -and $Master.Temperature -ne $SourceRecord.Temperature) {
                return $false
            }
            return $true
        }
    }

    return $false
}

function Get-AsiToPixRequiredMasterMetadataIssue {
    param(
        [Parameter(Mandatory = $true)]
        [object]$Master
    )

    switch ($Master.MasterType) {
        "Bias" {
            if ($null -eq $Master.Gain -or $null -eq $Master.Temperature) {
                return "The WBPP bias master name must contain GAIN and TEMP tags."
            }
        }
        "Dark" {
            if ($null -eq $Master.Gain -or $null -eq $Master.Temperature -or $null -eq $Master.Exposure) {
                return "The WBPP dark master name must contain GAIN, TEMP, and EXP or EXPOSURE tags."
            }
        }
        "Flat" {
            if ([string]::IsNullOrWhiteSpace([string]$Master.Filter)) {
                return "The WBPP flat master name must contain a FILTER tag."
            }
        }
    }

    return $null
}

function Get-AsiToPixDestinationMasterState {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$DestinationPath,

        [Parameter(Mandatory = $true)]
        [ValidateSet("Bias", "Dark", "Flat")]
        [string]$ExpectedMasterType
    )

    $normalizedDestinationPath = [System.IO.Path]::GetFullPath($DestinationPath)
    $destinationFolder = Split-Path -Path $normalizedDestinationPath -Parent
    if (Test-Path -LiteralPath $destinationFolder) {
        $destinationFolderItem = Get-Item -LiteralPath $destinationFolder -Force -ErrorAction Stop
        if (-not $destinationFolderItem.PSIsContainer) {
            return [PSCustomObject]@{
                State              = "Conflict"
                ExistingMasterPath = $null
                ExistingMasterInfo = $null
                ExistingMasterCount = 0
                Reason             = "The destination folder path is occupied by a file: '$destinationFolder'."
            }
        }
    } else {
        return [PSCustomObject]@{
            State               = "None"
            ExistingMasterPath  = $null
            ExistingMasterInfo  = $null
            ExistingMasterCount = 0
            Reason              = $null
        }
    }

    if (Test-Path -LiteralPath $normalizedDestinationPath -PathType Container) {
        return [PSCustomObject]@{
            State               = "Conflict"
            ExistingMasterPath  = $null
            ExistingMasterInfo  = $null
            ExistingMasterCount = 0
            Reason              = "A directory exists at the planned master file path: '$normalizedDestinationPath'."
        }
    }

    $masterFiles = [System.Collections.Generic.List[object]]::new()
    foreach ($file in @(Get-ChildItem -LiteralPath $destinationFolder -File -Filter "*.xisf" -ErrorAction Stop)) {
        $masterInfo = Get-AsiToPixWbppMasterInfo -FileName $file.Name
        if ($null -ne $masterInfo) {
            $masterFiles.Add([PSCustomObject]@{
                File = $file
                Info = $masterInfo
            })
        }
    }

    if ($masterFiles.Count -eq 0) {
        return [PSCustomObject]@{
            State               = "None"
            ExistingMasterPath  = $null
            ExistingMasterInfo  = $null
            ExistingMasterCount = 0
            Reason              = $null
        }
    }

    if ($masterFiles.Count -gt 1) {
        $existingPaths = @($masterFiles | ForEach-Object { $_.File.FullName } | Sort-Object)
        return [PSCustomObject]@{
            State               = "Conflict"
            ExistingMasterPath  = $null
            ExistingMasterInfo  = $null
            ExistingMasterCount = $masterFiles.Count
            Reason              = "The destination folder contains multiple masters: $($existingPaths -join '; ')."
        }
    }

    $existing = $masterFiles[0]
    $existingPath = [System.IO.Path]::GetFullPath($existing.File.FullName)
    if ($existingPath.Equals($normalizedDestinationPath, [System.StringComparison]::OrdinalIgnoreCase)) {
        return [PSCustomObject]@{
            State               = "Exact"
            ExistingMasterPath  = $existingPath
            ExistingMasterInfo  = $existing.Info
            ExistingMasterCount = 1
            Reason              = "A master with the canonical name already exists."
        }
    }

    $expectedFileName = Split-Path -Path $normalizedDestinationPath -Leaf
    $isLegacyName = $existing.Info.MasterType -eq $ExpectedMasterType -and
        $existing.Info.CleanFileName.Equals($expectedFileName, [System.StringComparison]::OrdinalIgnoreCase)
    if ($isLegacyName) {
        return [PSCustomObject]@{
            State               = "Legacy"
            ExistingMasterPath  = $existingPath
            ExistingMasterInfo  = $existing.Info
            ExistingMasterCount = 1
            Reason              = "The existing master has a legacy WBPP metadata name."
        }
    }

    return [PSCustomObject]@{
        State               = "Other"
        ExistingMasterPath  = $existingPath
        ExistingMasterInfo  = $existing.Info
        ExistingMasterCount = 1
        Reason              = "The destination folder contains one master with a different canonical identity."
    }
}

function Set-AsiToPixMasterExportEntryDestinationState {
    [CmdletBinding()]
    [OutputType([void])]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Entry,

        [Parameter(Mandatory = $true)]
        [object]$DestinationState
    )

    $Entry.ExistingMasterPath = $DestinationState.ExistingMasterPath
    $Entry.ExistingMasterCount = $DestinationState.ExistingMasterCount
    $Entry.Reason = $DestinationState.Reason
    switch ($DestinationState.State) {
        "None" {
            $Entry.Status = "Planned"
            $Entry.Reason = "The destination folder does not contain a master."
        }
        "Exact" {
            $Entry.Status = "Exists"
        }
        "Legacy" {
            $Entry.Status = "Exists"
        }
        "Other" {
            $Entry.Status = "Exists"
        }
        "Conflict" {
            $Entry.Status = "Conflict"
        }
        default {
            throw "Unknown destination master state '$($DestinationState.State)' for '$($Entry.DestinationPath)'."
        }
    }
}

function Read-AsiToPixMasterExportConfirmation {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Prompt,

        [Parameter(Mandatory = $true)]
        [scriptblock]$InputReader
    )

    while ($true) {
        $answerValue = & $InputReader "$Prompt [Y/N]"
        if ($null -eq $answerValue) {
            throw "Confirmation input ended before an answer was received."
        }

        $answer = ([string]$answerValue).Trim()
        $russianYes = [string][char]0x0434
        $russianNo = [string][char]0x043D
        if ($answer -ieq "y" -or $answer -ieq $russianYes) {
            return $true
        }
        if ($answer -ieq "n" -or $answer -ieq $russianNo) {
            return $false
        }

        Write-Host "Please answer Y/N or Russian yes/no initials." -ForegroundColor Yellow
    }
}

function Read-AsiToPixMasterExportSelection {
    [CmdletBinding()]
    [OutputType([object])]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$ChoiceCandidate,

        [Parameter(Mandatory = $true)]
        [scriptblock]$InputReader
    )

    if ($ChoiceCandidate.Count -lt 2) {
        throw "A master selection must contain at least two candidate files."
    }

    while ($true) {
        $answerValue = & $InputReader "Choose source [1-$($ChoiceCandidate.Count), S to skip]"
        if ($null -eq $answerValue) {
            throw "Selection input ended before an answer was received."
        }

        $answer = ([string]$answerValue).Trim()
        if ($answer -ieq "s" -or $answer -ieq "skip") {
            return $null
        }

        $choiceIndex = 0
        if ([int]::TryParse($answer, [ref]$choiceIndex) -and
            $choiceIndex -ge 1 -and $choiceIndex -le $ChoiceCandidate.Count) {
            return $ChoiceCandidate[$choiceIndex - 1]
        }

        Write-Host "Enter a number from 1 to $($ChoiceCandidate.Count), or S to skip." -ForegroundColor Yellow
    }
}

function Copy-AsiToPixMasterFile {
    [CmdletBinding(SupportsShouldProcess = $true)]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,

        [Parameter(Mandatory = $true)]
        [string]$DestinationPath
    )

    $normalizedSourcePath = (Resolve-Path -LiteralPath $SourcePath -ErrorAction Stop).ProviderPath
    $normalizedDestinationPath = [System.IO.Path]::GetFullPath($DestinationPath)
    if (Test-Path -LiteralPath $normalizedDestinationPath) {
        throw "Refusing to copy because the destination path already exists: '$normalizedDestinationPath'."
    }

    if (-not $PSCmdlet.ShouldProcess($normalizedDestinationPath, "Copy master from '$normalizedSourcePath'")) {
        return $false
    }

    $destinationFolder = Split-Path -Path $normalizedDestinationPath -Parent
    if (-not (Test-Path -LiteralPath $destinationFolder -PathType Container)) {
        New-Item -ItemType Directory -Path $destinationFolder -Force -ErrorAction Stop | Out-Null
    }

    [System.IO.File]::Copy($normalizedSourcePath, $normalizedDestinationPath, $false)
    $sourceLength = (Get-Item -LiteralPath $normalizedSourcePath -ErrorAction Stop).Length
    $destinationLength = (Get-Item -LiteralPath $normalizedDestinationPath -ErrorAction Stop).Length
    if ($destinationLength -ne $sourceLength) {
        Remove-Item -LiteralPath $normalizedDestinationPath -Force -ErrorAction Stop
        throw "Copied master size mismatch for '$normalizedDestinationPath'. The incomplete copy was removed."
    }

    return $true
}

function Get-AsiToPixMasterExportPlan {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Metadata,

        [AllowNull()]
        [AllowEmptyString()]
        [string]$AstroPhotoRoot,

        [AllowNull()]
        [AllowEmptyCollection()]
        [object[]]$SourceRecord
    )

    Assert-AsiToPixProjectMetadata -Metadata $Metadata -MetadataPath "<in-memory project metadata>"
    $pixPath = [System.IO.Path]::GetFullPath([string](Get-AsiToPixObjectPropertyValue -InputObject $Metadata -Name "PixPath"))
    if (-not (Test-Path -LiteralPath $pixPath -PathType Container)) {
        throw "PixInsight project folder not found: '$pixPath'."
    }

    $configuredMasterPath = [string](Get-AsiToPixObjectPropertyValue -InputObject $Metadata -Name "WbppMasterPath")
    $masterPath = if ([string]::IsNullOrWhiteSpace($configuredMasterPath)) {
        Join-Path -Path $pixPath -ChildPath "master"
    } else {
        if (-not [System.IO.Path]::IsPathRooted($configuredMasterPath)) {
            throw "WbppMasterPath in project metadata must be absolute: '$configuredMasterPath'."
        }
        [System.IO.Path]::GetFullPath($configuredMasterPath)
    }
    if (-not (Test-Path -LiteralPath $masterPath -PathType Container)) {
        throw "WBPP master folder not found: '$masterPath'."
    }

    $diagnostics = [System.Collections.Generic.List[string]]::new()
    $cameras = @(Get-AsiToPixObjectPropertyValue -InputObject $Metadata -Name "Cameras")
    $projectRoot = Split-Path -Path $pixPath -Parent
    $configuredProjectSourcePath = [string](Get-AsiToPixObjectPropertyValue -InputObject $Metadata -Name "ProjectSourcePath")
    $projectSourcePath = if ([string]::IsNullOrWhiteSpace($configuredProjectSourcePath)) {
        Join-Path -Path $projectRoot -ChildPath "Source"
    } else {
        if (-not [System.IO.Path]::IsPathRooted($configuredProjectSourcePath)) {
            throw "ProjectSourcePath in project metadata must be absolute: '$configuredProjectSourcePath'."
        }
        [System.IO.Path]::GetFullPath($configuredProjectSourcePath)
    }
    $metadataSourceProperty = $Metadata.PSObject.Properties["CalibrationSources"]
    $hasMetadataSources = $null -ne $metadataSourceProperty
    $requiresProjectSourceLinks = -not $PSBoundParameters.ContainsKey("SourceRecord") -and -not $hasMetadataSources
    if ($requiresProjectSourceLinks -and -not (Test-Path -LiteralPath $projectSourcePath -PathType Container)) {
        throw "Project source folder not found: '$projectSourcePath'."
    }

    $sourceRecords = @(if ($PSBoundParameters.ContainsKey("SourceRecord")) {
        $SourceRecord
    } elseif ($hasMetadataSources) {
        Get-AsiToPixMetadataCalibrationSourceRecord `
            -MetadataSource @($metadataSourceProperty.Value) `
            -CameraMetadata $cameras `
            -AstroPhotoRoot $AstroPhotoRoot `
            -Diagnostic $diagnostics
    } else {
        Get-AsiToPixProjectSourceRecord `
            -ProjectSourcePath $projectSourcePath `
            -CameraMetadata $cameras `
            -AstroPhotoRoot $AstroPhotoRoot `
            -Diagnostic $diagnostics
    })

    $allXisfFiles = @(Get-ChildItem -LiteralPath $masterPath -File -Filter "*.xisf" -ErrorAction Stop)
    $entries = [System.Collections.Generic.List[object]]::new()
    $candidates = [System.Collections.Generic.List[object]]::new()
    $eligibleMasterCount = 0

    foreach ($file in $allXisfFiles) {
        $master = Get-AsiToPixWbppMasterInfo -FileName $file.Name
        if ($null -eq $master) {
            continue
        }
        $eligibleMasterCount++

        $requiredIssue = Get-AsiToPixRequiredMasterMetadataIssue -Master $master
        $allIssues = @($master.Issues)
        if ($null -ne $requiredIssue) {
            $allIssues += $requiredIssue
        }
        if ($allIssues.Count -gt 0) {
            $entries.Add([PSCustomObject]@{
                Status          = "Skipped"
                MasterType      = $master.MasterType
                SourcePath      = $file.FullName
                DestinationPath = $null
                DuplicateOf     = $null
                Reason          = $allIssues -join " "
            })
            continue
        }

        $sourceMatches = @($sourceRecords | Where-Object { Test-AsiToPixMasterSourceMatch -Master $master -SourceRecord $_ })
        $destinationFolders = @($sourceMatches.DestinationFolder | Sort-Object -Unique)
        if ($destinationFolders.Count -eq 0) {
            $entries.Add([PSCustomObject]@{
                Status          = "Skipped"
                MasterType      = $master.MasterType
                SourcePath      = $file.FullName
                DestinationPath = $null
                DuplicateOf     = $null
                Reason          = "No matching calibration source mapping was found."
            })
            continue
        }
        if ($destinationFolders.Count -gt 1) {
            $entries.Add([PSCustomObject]@{
                Status          = "Conflict"
                MasterType      = $master.MasterType
                SourcePath      = $file.FullName
                DestinationPath = $null
                DuplicateOf     = $null
                Reason          = "Matching Source links point to multiple destination folders: $($destinationFolders -join '; ')."
            })
            continue
        }

        $candidates.Add([PSCustomObject]@{
            Status               = "Candidate"
            MasterType           = $master.MasterType
            SourcePath           = $file.FullName
            SourceLength         = $file.Length
            DestinationPath      = Join-Path -Path $destinationFolders[0] -ChildPath $master.CleanFileName
            ExistingMasterPath   = $null
            ExistingMasterCount  = 0
            DuplicateOf          = $null
            ChoiceCandidates     = @()
            Reason               = $null
        })
    }

    $candidateGroups = [System.Collections.Generic.Dictionary[string, System.Collections.Generic.List[object]]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    foreach ($candidate in $candidates) {
        if (-not $candidateGroups.ContainsKey($candidate.DestinationPath)) {
            $candidateGroups.Add($candidate.DestinationPath, [System.Collections.Generic.List[object]]::new())
        }
        $candidateGroups[$candidate.DestinationPath].Add($candidate)
    }

    foreach ($destinationPath in @($candidateGroups.Keys | Sort-Object)) {
        $group = @($candidateGroups[$destinationPath] | Sort-Object SourcePath)
        $distinctLengths = @($group.SourceLength | Sort-Object -Unique)
        if ($distinctLengths.Count -gt 1) {
            $selection = $group[0]
            $destinationState = Get-AsiToPixDestinationMasterState `
                -DestinationPath $destinationPath `
                -ExpectedMasterType $selection.MasterType
            if ($destinationState.State -ne "None") {
                Set-AsiToPixMasterExportEntryDestinationState `
                    -Entry $selection `
                    -DestinationState $destinationState
                $entries.Add($selection)
                continue
            }
            $selection.Status = "Conflict"
            $selection.ChoiceCandidates = @($group)
            $selection.Reason = "Logical duplicates have different file sizes. Select one source before executing the export plan."
            $entries.Add($selection)
            continue
        }

        $primary = $group[0]
        $destinationState = Get-AsiToPixDestinationMasterState `
            -DestinationPath $destinationPath `
            -ExpectedMasterType $primary.MasterType
        Set-AsiToPixMasterExportEntryDestinationState `
            -Entry $primary `
            -DestinationState $destinationState
        $entries.Add($primary)

        foreach ($duplicate in @($group | Select-Object -Skip 1)) {
            $duplicate.Status = "Duplicate"
            $duplicate.DuplicateOf = $primary.SourcePath
            $duplicate.Reason = "Same logical destination and file size; this duplicate would be omitted."
            $entries.Add($duplicate)
        }
    }

    return [PSCustomObject]@{
        MasterPath          = $masterPath
        ProjectSourcePath   = $projectSourcePath
        XisfFileCount       = $allXisfFiles.Count
        EligibleMasterCount = $eligibleMasterCount
        IgnoredFileCount    = $allXisfFiles.Count - $eligibleMasterCount
        SourceRecordCount   = $sourceRecords.Count
        Diagnostics         = @($diagnostics)
        Entries             = @($entries | Sort-Object MasterType, DestinationPath, SourcePath)
    }
}

function ConvertTo-AsiToPixMasterExportFileSizeText {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [long]$ByteCount
    )

    $units = @("B", "KiB", "MiB", "GiB", "TiB")
    $size = [double]$ByteCount
    $unitIndex = 0
    while ($size -ge 1024 -and $unitIndex -lt ($units.Count - 1)) {
        $size = $size / 1024
        $unitIndex++
    }

    $invariantCulture = [System.Globalization.CultureInfo]::InvariantCulture
    return "$($size.ToString('0.0', $invariantCulture)) $($units[$unitIndex])"
}

function Get-AsiToPixMasterExportDestinationDisplay {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Entry
    )

    $destinationPath = [string](Get-AsiToPixObjectPropertyValue -InputObject $Entry -Name "DestinationPath")
    $masterType = [string](Get-AsiToPixObjectPropertyValue -InputObject $Entry -Name "MasterType")
    if ([string]::IsNullOrWhiteSpace($destinationPath)) {
        return [PSCustomObject]@{
            Camera          = "Unmapped"
            Category        = "Unmapped"
            DestinationRoot = $null
            Setup           = $null
            LeafDestination = $null
        }
    }

    $normalizedDestinationPath = [System.IO.Path]::GetFullPath($destinationPath)
    $destinationFolder = Split-Path -Path $normalizedDestinationPath -Parent
    $categoryRoot = $null
    $masterContainer = $null
    $currentFolder = $destinationFolder
    while (-not [string]::IsNullOrWhiteSpace($currentFolder)) {
        $possibleMasterContainer = Split-Path -Path $currentFolder -Parent
        if ([string]::IsNullOrWhiteSpace($possibleMasterContainer)) {
            break
        }
        if ((Split-Path -Path $possibleMasterContainer -Leaf) -ieq "Master") {
            $categoryRoot = $currentFolder
            $masterContainer = $possibleMasterContainer
            break
        }

        $parentFolder = Split-Path -Path $currentFolder -Parent
        if ([string]::IsNullOrWhiteSpace($parentFolder) -or
            $parentFolder.Equals($currentFolder, [System.StringComparison]::OrdinalIgnoreCase)) {
            break
        }
        $currentFolder = $parentFolder
    }

    if ($null -eq $categoryRoot) {
        $fallbackCategory = switch ($masterType) {
            "Bias" { "Biases" }
            "Dark" { "Darks" }
            "Flat" { "Flats" }
            default { "Masters" }
        }
        return [PSCustomObject]@{
            Camera          = "Unknown camera"
            Category        = $fallbackCategory
            DestinationRoot = $destinationFolder
            Setup           = $null
            LeafDestination = Split-Path -Path $normalizedDestinationPath -Leaf
        }
    }

    $category = switch ((Split-Path -Path $categoryRoot -Leaf).ToLowerInvariant()) {
        "bias" { "Biases" }
        "biases" { "Biases" }
        "dark" { "Darks" }
        "darks" { "Darks" }
        "flat" { "Flats" }
        "flats" { "Flats" }
        "flat-dark" { "FlatDarks" }
        "flat-darks" { "FlatDarks" }
        "flatdark" { "FlatDarks" }
        "flatdarks" { "FlatDarks" }
        default { "Masters" }
    }
    $cameraFolder = Split-Path -Path $masterContainer -Parent
    $camera = Split-Path -Path $cameraFolder -Leaf
    $relativeDestination = Get-AsiToPixRelativeChildPath `
        -RootPath $categoryRoot `
        -ChildPath $normalizedDestinationPath
    if ([string]::IsNullOrWhiteSpace($relativeDestination)) {
        $relativeDestination = Split-Path -Path $normalizedDestinationPath -Leaf
    }

    $setup = $null
    $leafDestination = $relativeDestination
    if ($category -eq "Flats") {
        $relativeParts = @($relativeDestination -split '[\\/]')
        if ($relativeParts.Count -gt 1 -and -not [string]::IsNullOrWhiteSpace($relativeParts[0])) {
            $setup = $relativeParts[0]
            $leafDestination = ($relativeParts | Select-Object -Skip 1) -join "\"
        }
    }

    return [PSCustomObject]@{
        Camera          = $camera
        Category        = $category
        DestinationRoot = $categoryRoot
        Setup           = $setup
        LeafDestination = $leafDestination
    }
}

function Get-AsiToPixMasterExportTreeLeafText {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Entry,

        [Parameter(Mandatory = $true)]
        [object]$Display
    )

    $status = [string](Get-AsiToPixObjectPropertyValue -InputObject $Entry -Name "Status")
    $sourcePath = [string](Get-AsiToPixObjectPropertyValue -InputObject $Entry -Name "SourcePath")
    $sourceName = if ([string]::IsNullOrWhiteSpace($sourcePath)) {
        "<unknown source>"
    } else {
        Split-Path -Path $sourcePath -Leaf
    }
    $destination = [string](Get-AsiToPixObjectPropertyValue -InputObject $Display -Name "LeafDestination")
    if ([string]::IsNullOrWhiteSpace($destination)) {
        $destination = "<no destination>"
    }

    switch ($status) {
        "Planned" { return "[Copy] $sourceName --> $destination" }
        "Exists" {
            $existingPath = [string](Get-AsiToPixObjectPropertyValue -InputObject $Entry -Name "ExistingMasterPath")
            $existingName = if ([string]::IsNullOrWhiteSpace($existingPath)) {
                Split-Path -Path $destination -Leaf
            } else {
                Split-Path -Path $existingPath -Leaf
            }
            return "[Exists] $existingName --> $destination"
        }
        "Legacy" { return "[Exists] $sourceName --> $destination" }
        "ExistingOther" { return "[Exists] $sourceName --> $destination" }
        "Duplicate" { return "[Duplicate omitted] $sourceName --> $destination" }
        "Conflict" {
            $choiceCandidates = @(Get-AsiToPixObjectPropertyValue -InputObject $Entry -Name "ChoiceCandidates")
            if ($choiceCandidates.Count -gt 1) {
                return "[Conflict] choose one of $($choiceCandidates.Count) --> $destination"
            }
            return "[Conflict] $sourceName --> $destination"
        }
        "Skipped" { return "[Skipped] $sourceName" }
        default { return "[$status] $sourceName --> $destination" }
    }
}

function Get-AsiToPixMasterExportTreeLeafColor {
    [CmdletBinding()]
    [OutputType([System.ConsoleColor])]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Entry
    )

    switch ([string](Get-AsiToPixObjectPropertyValue -InputObject $Entry -Name "Status")) {
        "Planned" { return [System.ConsoleColor]::Green }
        "Exists" { return [System.ConsoleColor]::Yellow }
        "Legacy" { return [System.ConsoleColor]::Yellow }
        "ExistingOther" { return [System.ConsoleColor]::Yellow }
        "Duplicate" { return [System.ConsoleColor]::DarkYellow }
        "Conflict" { return [System.ConsoleColor]::Red }
        "Skipped" { return [System.ConsoleColor]::Yellow }
        default { return [System.ConsoleColor]::Gray }
    }
}

function Write-AsiToPixMasterExportTreeLeaf {
    [CmdletBinding()]
    [OutputType([void])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Indent,

        [Parameter(Mandatory = $true)]
        [object]$Entry,

        [Parameter(Mandatory = $true)]
        [object]$Display
    )

    $text = Get-AsiToPixMasterExportTreeLeafText -Entry $Entry -Display $Display
    $color = Get-AsiToPixMasterExportTreeLeafColor -Entry $Entry
    Write-Host "$Indent $text" -ForegroundColor $color

    $choiceCandidates = @(Get-AsiToPixObjectPropertyValue -InputObject $Entry -Name "ChoiceCandidates")
    if ([string](Get-AsiToPixObjectPropertyValue -InputObject $Entry -Name "Status") -ne "Conflict" -or
        $choiceCandidates.Count -lt 2) {
        return
    }

    $choiceIndent = $Indent.Substring(0, $Indent.Length - 2) + "|  |-"
    $choiceIndex = 1
    foreach ($candidate in $choiceCandidates) {
        $candidatePath = [string](Get-AsiToPixObjectPropertyValue -InputObject $candidate -Name "SourcePath")
        $candidateName = Split-Path -Path $candidatePath -Leaf
        $candidateLength = [long](Get-AsiToPixObjectPropertyValue -InputObject $candidate -Name "SourceLength")
        $sizeText = ConvertTo-AsiToPixMasterExportFileSizeText -ByteCount $candidateLength
        Write-Host "$choiceIndent [$choiceIndex] $candidateName ($sizeText)" -ForegroundColor DarkYellow
        $choiceIndex++
    }
}

function Write-AsiToPixMasterExportPlan {
    [CmdletBinding()]
    [OutputType([void])]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Plan
    )

    Write-Host "`n[Plan] calibration masters from: $($Plan.MasterPath)" -ForegroundColor Cyan
    Write-Host "       project sources: $($Plan.ProjectSourcePath)" -ForegroundColor DarkGray

    foreach ($diagnostic in @($Plan.Diagnostics)) {
        Write-Host "[Source warning] $diagnostic" -ForegroundColor Yellow
    }

    if ($Plan.EligibleMasterCount -eq 0) {
        Write-Host "[Info] No WBPP bias, dark, or flat masters were found in '$($Plan.MasterPath)'." -ForegroundColor Yellow
    }

    $displayEntries = foreach ($entry in @($Plan.Entries)) {
        [PSCustomObject]@{
            Entry   = $entry
            Display = Get-AsiToPixMasterExportDestinationDisplay -Entry $entry
        }
    }
    $mappedEntries = @($displayEntries | Where-Object {
        -not [string]::IsNullOrWhiteSpace([string]$_.Display.DestinationRoot)
    })

    $categoryOrder = @{ "Flats" = 0; "Darks" = 1; "FlatDarks" = 2; "Biases" = 3; "Masters" = 4 }
    foreach ($cameraGroup in @($mappedEntries | Group-Object -Property { $_.Display.Camera } | Sort-Object Name)) {
        Write-Host "`n[Camera] $($cameraGroup.Name)" -ForegroundColor Cyan
        $branches = @($cameraGroup.Group | Group-Object -Property {
            "$($_.Display.Category)`0$($_.Display.DestinationRoot)"
        })
        $sortedBranches = @($branches | Sort-Object `
            @{ Expression = {
                $category = $_.Group[0].Display.Category
                if ($categoryOrder.ContainsKey($category)) { $categoryOrder[$category] } else { 99 }
            } }, `
            Name)
        foreach ($branch in $sortedBranches) {
            $branchEntries = @($branch.Group | Sort-Object `
                @{ Expression = { $_.Display.Setup } }, `
                @{ Expression = { $_.Display.LeafDestination } }, `
                @{ Expression = { $_.Entry.SourcePath } })
            $firstDisplay = $branchEntries[0].Display
            Write-Host "  |- [$($firstDisplay.Category)] copy to: $($firstDisplay.DestinationRoot)" -ForegroundColor White

            if ($firstDisplay.Category -eq "Flats") {
                foreach ($setupGroup in @($branchEntries | Group-Object -Property {
                    if ([string]::IsNullOrWhiteSpace([string]$_.Display.Setup)) {
                        "<unknown setup>"
                    } else {
                        $_.Display.Setup
                    }
                } | Sort-Object Name)) {
                    Write-Host "  |  |- [$($setupGroup.Name)]" -ForegroundColor White
                    foreach ($displayEntry in @($setupGroup.Group | Sort-Object `
                        @{ Expression = { $_.Display.LeafDestination } }, `
                        @{ Expression = { $_.Entry.SourcePath } })) {
                        Write-AsiToPixMasterExportTreeLeaf `
                            -Indent "  |  |  |-" `
                            -Entry $displayEntry.Entry `
                            -Display $displayEntry.Display
                    }
                }
            } else {
                foreach ($displayEntry in $branchEntries) {
                    Write-AsiToPixMasterExportTreeLeaf `
                        -Indent "  |  |-" `
                        -Entry $displayEntry.Entry `
                        -Display $displayEntry.Display
                }
            }
        }
    }

    $unmappedEntries = @($displayEntries | Where-Object {
        [string]::IsNullOrWhiteSpace([string]$_.Display.DestinationRoot)
    })
    if ($unmappedEntries.Count -gt 0) {
        Write-Host "`n[Unmapped]" -ForegroundColor Yellow
        foreach ($displayEntry in $unmappedEntries) {
            $entry = $displayEntry.Entry
            $sourcePath = [string](Get-AsiToPixObjectPropertyValue -InputObject $entry -Name "SourcePath")
            $sourceName = if ([string]::IsNullOrWhiteSpace($sourcePath)) {
                "<unknown source>"
            } else {
                Split-Path -Path $sourcePath -Leaf
            }
            $reason = [string](Get-AsiToPixObjectPropertyValue -InputObject $entry -Name "Reason")
            $color = Get-AsiToPixMasterExportTreeLeafColor -Entry $entry
            Write-Host "  |- ${sourceName}: $reason" -ForegroundColor $color
        }
    }

    $entries = @($Plan.Entries)
    $plannedCount = @($entries | Where-Object { $_.Status -eq "Planned" }).Count
    $existingCount = @($entries | Where-Object { $_.Status -eq "Exists" }).Count
    $duplicateCount = @($entries | Where-Object { $_.Status -eq "Duplicate" }).Count
    $skippedCount = @($entries | Where-Object { $_.Status -eq "Skipped" }).Count
    $conflictCount = @($entries | Where-Object { $_.Status -eq "Conflict" }).Count

    Write-Host "`n[Summary]" -ForegroundColor Cyan
    Write-Host "  WBPP .xisf files    : $($Plan.XisfFileCount)"
    Write-Host "  Calibration masters : $($Plan.EligibleMasterCount)"
    Write-Host "  Planned copies      : $plannedCount" -ForegroundColor Green
    Write-Host "  Already exist       : $existingCount" -ForegroundColor Yellow
    Write-Host "  Duplicates omitted  : $duplicateCount" -ForegroundColor DarkYellow
    Write-Host "  Skipped / conflicts : $skippedCount / $conflictCount"
    Write-Host "  Ignored non-masters : $($Plan.IgnoredFileCount)" -ForegroundColor DarkGray
    Write-Host "`nNo changes have been made yet." -ForegroundColor Cyan
}

function Invoke-AsiToPixMasterExportPlan {
    [CmdletBinding(SupportsShouldProcess = $true)]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory = $true)]
        [object]$Plan,

        [scriptblock]$ConfirmationReader = {
            param($Prompt)
            Read-Host $Prompt
        },

        [scriptblock]$SelectionReader = {
            param($Prompt)
            Read-Host $Prompt
        }
    )

    $copiedCount = 0
    $declinedCount = 0
    $blockedCount = 0
    $existingCount = 0
    $whatIfCount = 0
    $selectedCount = 0
    $selectionSkippedCount = 0
    $conflictPendingCount = 0
    $selectionChanged = $false
    $resolvedEntries = [System.Collections.Generic.List[object]]::new()

    Write-Host "`n--- Resolve master conflicts ---" -ForegroundColor Cyan
    foreach ($entry in @($Plan.Entries)) {
        $choiceCandidates = @(Get-AsiToPixObjectPropertyValue -InputObject $entry -Name "ChoiceCandidates")
        $requiresSelection = $entry.Status -eq "Conflict" -and $choiceCandidates.Count -gt 1
        if (-not $requiresSelection) {
            $resolvedEntries.Add($entry)
            continue
        }

        $destinationPath = [string](Get-AsiToPixObjectPropertyValue -InputObject $entry -Name "DestinationPath")
        $destinationState = Get-AsiToPixDestinationMasterState `
            -DestinationPath $destinationPath `
            -ExpectedMasterType $entry.MasterType
        if ($destinationState.State -ne "None") {
            $entry.ChoiceCandidates = @()
            Set-AsiToPixMasterExportEntryDestinationState `
                -Entry $entry `
                -DestinationState $destinationState
            $resolvedEntries.Add($entry)
            $selectionChanged = $true
            continue
        }

        if ($WhatIfPreference) {
            Write-Host "`n[WHATIF Conflict] $destinationPath" -ForegroundColor DarkYellow
            Write-Host "  A source selection is required; WhatIf does not prompt or choose a candidate." -ForegroundColor DarkGray
            $resolvedEntries.Add($entry)
            $conflictPendingCount++
            continue
        }

        Write-Host "`n[Conflict] select one master for: $destinationPath" -ForegroundColor Red
        $choiceIndex = 1
        foreach ($choiceCandidate in $choiceCandidates) {
            $choicePath = [string](Get-AsiToPixObjectPropertyValue -InputObject $choiceCandidate -Name "SourcePath")
            $choiceName = Split-Path -Path $choicePath -Leaf
            $choiceLength = [long](Get-AsiToPixObjectPropertyValue -InputObject $choiceCandidate -Name "SourceLength")
            $choiceSize = ConvertTo-AsiToPixMasterExportFileSizeText -ByteCount $choiceLength
            Write-Host "  [$choiceIndex] $choiceName ($choiceSize)" -ForegroundColor Gray
            $choiceIndex++
        }

        $selectedCandidate = Read-AsiToPixMasterExportSelection `
            -ChoiceCandidate $choiceCandidates `
            -InputReader $SelectionReader
        if ($null -eq $selectedCandidate) {
            $entry.Status = "Skipped"
            $entry.ChoiceCandidates = @()
            $entry.Reason = "The user skipped source selection for this logical duplicate group."
            $resolvedEntries.Add($entry)
            $selectionSkippedCount++
            $selectionChanged = $true
            continue
        }

        $selectedDestinationPath = [string](Get-AsiToPixObjectPropertyValue `
            -InputObject $selectedCandidate `
            -Name "DestinationPath")
        if (-not $selectedDestinationPath.Equals(
            $destinationPath,
            [System.StringComparison]::OrdinalIgnoreCase
        )) {
            throw "Selected master destination differs from its logical duplicate group: '$selectedDestinationPath' vs '$destinationPath'."
        }

        $selectedCandidate.ChoiceCandidates = @()
        $selectedDestinationState = Get-AsiToPixDestinationMasterState `
            -DestinationPath $selectedCandidate.DestinationPath `
            -ExpectedMasterType $selectedCandidate.MasterType
        Set-AsiToPixMasterExportEntryDestinationState `
            -Entry $selectedCandidate `
            -DestinationState $selectedDestinationState
        $resolvedEntries.Add($selectedCandidate)
        $selectedCount++
        $selectionChanged = $true
    }

    if ($selectionChanged) {
        $Plan.Entries = @($resolvedEntries | Sort-Object MasterType, DestinationPath, SourcePath)
        Write-Host "`n--- Updated master export plan ---" -ForegroundColor Cyan
        Write-AsiToPixMasterExportPlan -Plan $Plan
    }

    $actionableEntries = @($Plan.Entries | Where-Object { $_.Status -eq "Planned" })
    if ($actionableEntries.Count -eq 0) {
        Write-Host "`n[Info] No planned master copies to execute." -ForegroundColor Yellow
    } elseif ($WhatIfPreference) {
        foreach ($entry in $actionableEntries) {
            Copy-AsiToPixMasterFile `
                -SourcePath $entry.SourcePath `
                -DestinationPath $entry.DestinationPath `
                -WhatIf `
                -Confirm:$false | Out-Null
            $whatIfCount++
        }
    } else {
        $executeConfirmed = Read-AsiToPixMasterExportConfirmation `
            -Prompt "Execute this plan ($($actionableEntries.Count) master copy/copies)?" `
            -InputReader $ConfirmationReader
        if (-not $executeConfirmed) {
            $declinedCount = $actionableEntries.Count
            Write-Host "[Info] Plan execution declined; no files were copied." -ForegroundColor DarkYellow
        } else {
            foreach ($entry in $actionableEntries) {
                if (-not (Test-Path -LiteralPath $entry.SourcePath -PathType Leaf)) {
                    throw "WBPP master disappeared before export: '$($entry.SourcePath)'."
                }
                $currentSourceLength = (Get-Item -LiteralPath $entry.SourcePath -ErrorAction Stop).Length
                if ($currentSourceLength -ne $entry.SourceLength) {
                    throw "WBPP master changed after the plan was built: '$($entry.SourcePath)'. Build the export plan again."
                }

                $destinationState = Get-AsiToPixDestinationMasterState `
                    -DestinationPath $entry.DestinationPath `
                    -ExpectedMasterType $entry.MasterType
                if ($destinationState.State -ne "None") {
                    if ($destinationState.State -eq "Conflict") {
                        $blockedCount++
                        Write-Host "[Conflict] $($entry.DestinationPath) was not changed: $($destinationState.Reason)" -ForegroundColor Red
                    } else {
                        $existingCount++
                        Write-Host "[Exists] $($entry.DestinationPath) was not overwritten." -ForegroundColor Yellow
                    }
                    continue
                }

                $copied = Copy-AsiToPixMasterFile `
                    -SourcePath $entry.SourcePath `
                    -DestinationPath $entry.DestinationPath `
                    -Confirm:$false
                if ($copied) {
                    Write-Host "[Copied] $($entry.DestinationPath)" -ForegroundColor Green
                    $copiedCount++
                }
            }
        }
    }

    Write-Host "`n--- Apply summary ---" -ForegroundColor Cyan
    Write-Host "Copied             : $copiedCount" -ForegroundColor Green
    Write-Host "Already exists     : $existingCount" -ForegroundColor Yellow
    Write-Host "Declined           : $declinedCount" -ForegroundColor DarkYellow
    Write-Host "Blocked            : $blockedCount" -ForegroundColor Red
    Write-Host "WhatIf operations  : $whatIfCount" -ForegroundColor DarkGray
    Write-Host "Selections made    : $selectedCount" -ForegroundColor Green
    Write-Host "Selections skipped : $selectionSkippedCount" -ForegroundColor DarkYellow
    Write-Host "Conflicts pending  : $conflictPendingCount" -ForegroundColor DarkYellow

    return [PSCustomObject]@{
        CopiedCount           = $copiedCount
        ExistingCount         = $existingCount
        DeclinedCount         = $declinedCount
        BlockedCount          = $blockedCount
        WhatIfCount           = $whatIfCount
        SelectedCount         = $selectedCount
        SelectionSkippedCount = $selectionSkippedCount
        ConflictPendingCount  = $conflictPendingCount
    }
}

Export-ModuleMember -Function `
    Find-AsiToPixProcessingProjectMetadata, `
    Get-AsiToPixCleanMasterFileName, `
    Get-AsiToPixMasterExportPlan, `
    Get-AsiToPixWbppMasterInfo, `
    Invoke-AsiToPixMasterExportPlan, `
    Read-AsiToPixProjectMetadata, `
    Resolve-AsiToPixProjectMetadataPath, `
    Test-AsiToPixProjectMetadataNeedsAstroPhotoRoot, `
    Write-AsiToPixMasterExportPlan
