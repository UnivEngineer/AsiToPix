function Get-AsiToPixTotalityPhaseFolderName {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateSet("Totality", "C2", "C3")]
        [string]$Phase
    )

    switch ($Phase) {
        "Totality" { return "Totality" }
        "C2" { return "C2 beads" }
        "C3" { return "C3 beads" }
    }
}

function Get-AsiToPixTotalityExposureDirectory {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ProcessingRoot
    )

    if (-not (Test-Path -LiteralPath $ProcessingRoot -PathType Container)) {
        throw "Totality processing root not found: '$ProcessingRoot'."
    }

    $resolvedRoot = (Resolve-Path -LiteralPath $ProcessingRoot -ErrorAction Stop).ProviderPath
    return @(Get-ChildItem -LiteralPath $resolvedRoot -Directory -ErrorAction Stop |
        Where-Object { $_.Name -like "Total-*" } |
        Sort-Object -Property Name)
}

function Get-AsiToPixTotalityFrame {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [Alias("RegisteredPath")]
        [string]$FramePath,

        [switch]$AllowEmpty
    )

    if (-not (Test-Path -LiteralPath $FramePath -PathType Container)) {
        throw "Input frame folder not found: '$FramePath'."
    }

    $resolvedPath = (Resolve-Path -LiteralPath $FramePath -ErrorAction Stop).ProviderPath
    $namePattern = [regex]::new(
        "^(?<prefix>.+?_(?<exposure>\d+(?:\.\d+)?ms))_(?<index>\d+)(?:_.*)?\.(?<extension>xisf|fit|fits|fts)$",
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )
    $frames = [System.Collections.Generic.List[object]]::new()

    foreach ($file in @(Get-ChildItem -LiteralPath $resolvedPath -File -ErrorAction Stop)) {
        $match = $namePattern.Match($file.Name)
        if (-not $match.Success) {
            continue
        }

        $index = 0
        if (-not [int]::TryParse(
                $match.Groups["index"].Value,
                [System.Globalization.NumberStyles]::None,
                [System.Globalization.CultureInfo]::InvariantCulture,
                [ref]$index)) {
            throw "Frame index is outside the supported Int32 range: '$($file.FullName)'."
        }

        $frames.Add([PSCustomObject]@{
            Index     = $index
            IndexText = $match.Groups["index"].Value
            Name      = $file.Name
            Path      = $file.FullName
            Extension = $match.Groups["extension"].Value.ToLowerInvariant()
        })
    }

    if ($frames.Count -eq 0 -and -not $AllowEmpty) {
        throw "No processed *_<exposure>_<index>[optional suffix].xisf/.fit/.fits/.fts frames found in '$resolvedPath'."
    }

    $duplicates = @($frames | Group-Object -Property Index | Where-Object { $_.Count -gt 1 })
    if ($duplicates.Count -gt 0) {
        $duplicateDescription = @($duplicates | ForEach-Object {
            "index $($_.Name): $(@($_.Group.Path) -join ', ')"
        }) -join "; "
        throw "Duplicate input frame indexes found in '$resolvedPath': $duplicateDescription."
    }

    return @($frames | Sort-Object -Property Index, Name)
}

function Get-AsiToPixTotalityRawFrame {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$RawPath
    )

    if (-not (Test-Path -LiteralPath $RawPath -PathType Container)) {
        throw "Raw frame folder not found: '$RawPath'."
    }

    $resolvedPath = (Resolve-Path -LiteralPath $RawPath -ErrorAction Stop).ProviderPath
    $namePattern = [regex]::new(
        "^(?<prefix>.+)_(?<index>\d+)\.(?<extension>xisf|fit|fits|fts)$",
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )
    $frames = [System.Collections.Generic.List[object]]::new()

    foreach ($file in @(Get-ChildItem -LiteralPath $resolvedPath -File -ErrorAction Stop)) {
        $match = $namePattern.Match($file.Name)
        if (-not $match.Success) {
            continue
        }

        $index = 0
        if (-not [int]::TryParse(
                $match.Groups["index"].Value,
                [System.Globalization.NumberStyles]::None,
                [System.Globalization.CultureInfo]::InvariantCulture,
                [ref]$index)) {
            throw "Raw frame index is outside the supported Int32 range: '$($file.FullName)'."
        }

        $frames.Add([PSCustomObject]@{
            Index     = $index
            IndexText = $match.Groups["index"].Value
            Name      = $file.Name
            Path      = $file.FullName
            Extension = $match.Groups["extension"].Value.ToLowerInvariant()
        })
    }

    if ($frames.Count -eq 0) {
        throw "No raw *_<index>.xisf/.fit/.fits/.fts frames found in '$resolvedPath'."
    }

    $duplicates = @($frames | Group-Object -Property Index | Where-Object { $_.Count -gt 1 })
    if ($duplicates.Count -gt 0) {
        throw "Duplicate raw frame indexes found in '$resolvedPath': $(@($duplicates.Name) -join ', ')."
    }

    return @($frames | Sort-Object -Property Index, Name)
}

function Format-AsiToPixTotalityIndexList {
    [CmdletBinding()]
    param(
        [AllowEmptyCollection()]
        [int[]]$Indexes
    )

    $sortedIndexes = @($Indexes | Sort-Object -Unique)
    if ($sortedIndexes.Count -eq 0) {
        return "-"
    }

    $parts = [System.Collections.Generic.List[string]]::new()
    $rangeStart = $sortedIndexes[0]
    $rangeEnd = $rangeStart
    for ($offset = 1; $offset -lt $sortedIndexes.Count; $offset++) {
        $index = $sortedIndexes[$offset]
        if ($index -eq ($rangeEnd + 1)) {
            $rangeEnd = $index
            continue
        }

        $parts.Add($(if ($rangeStart -eq $rangeEnd) { "$rangeStart" } else { "$rangeStart-$rangeEnd" }))
        $rangeStart = $index
        $rangeEnd = $index
    }
    $parts.Add($(if ($rangeStart -eq $rangeEnd) { "$rangeStart" } else { "$rangeStart-$rangeEnd" }))
    return @($parts) -join ","
}

function Resolve-AsiToPixTotalityFrameManifest {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory)]
        [string]$ManifestPath,

        [Parameter(Mandatory)]
        [int]$MinimumIndex,

        [Parameter(Mandatory)]
        [int]$MaximumIndex,

        [scriptblock]$InputReader = { param($Message) Read-Host $Message }
    )

    if ($MinimumIndex -gt $MaximumIndex) {
        throw "Invalid raw frame index range: $MinimumIndex-$MaximumIndex."
    }
    $manifestInputReader = $InputReader

    function Read-ManifestInteger {
        param(
            [string]$Prompt,
            [int]$Minimum,
            [int]$Maximum
        )

        while ($true) {
            $answer = ([string](& $manifestInputReader $Prompt)).Trim()
            $value = 0
            if ([int]::TryParse($answer, [ref]$value) -and $value -ge $Minimum -and $value -le $Maximum) {
                return $value
            }
            Write-Host "Enter an integer from $Minimum through $Maximum." -ForegroundColor Yellow
        }
    }

    function Read-ManifestBoundaryIndex {
        param(
            [string]$Prompt,
            [int]$Minimum,
            [int]$Maximum
        )

        while ($true) {
            $answer = ([string](& $manifestInputReader $Prompt)).Trim()
            if ([string]::IsNullOrEmpty($answer)) {
                return $null
            }

            $value = 0
            if ([int]::TryParse($answer, [ref]$value) -and $value -ge $Minimum -and $value -le $Maximum) {
                return $value
            }
            Write-Host "Enter an integer from $Minimum through $Maximum, or press Enter if this boundary is absent." -ForegroundColor Yellow
        }
    }

    function ConvertTo-ManifestBoundaryIndex {
        param(
            [string]$PropertyName,
            [AllowNull()]
            [object]$Value
        )

        if ($null -eq $Value) {
            return $null
        }

        $convertedValue = 0
        $displayValue = [Convert]::ToString($Value, [System.Globalization.CultureInfo]::InvariantCulture)
        $isInteger = $Value -isnot [bool] -and [int]::TryParse(
            $displayValue,
            [System.Globalization.NumberStyles]::Integer,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [ref]$convertedValue
        )
        if (-not $isInteger -or $convertedValue -lt $MinimumIndex -or $convertedValue -gt $MaximumIndex) {
            $nullExample = '"{0}": null' -f $PropertyName
            $contactName = $PropertyName -replace "Index$", ""
            throw "Suspicious index for ${PropertyName}: $displayValue. Expected null or an integer in raw frame range $MinimumIndex-${MaximumIndex}. Please type $nullExample if your dataset has no $contactName contact frame. Manifest: '$ManifestPath'."
        }

        return $convertedValue
    }

    function Read-ManifestIndexList {
        param(
            [string]$Prompt,
            [int]$Minimum,
            [int]$Maximum
        )

        while ($true) {
            $answer = ([string](& $manifestInputReader $Prompt)).Trim()
            if ([string]::IsNullOrEmpty($answer)) {
                return [int[]]@()
            }

            try {
                $indexes = [System.Collections.Generic.List[int]]::new()
                foreach ($token in @($answer -split "[,;\s]+" | Where-Object { -not [string]::IsNullOrWhiteSpace($_) })) {
                    if ($token -match "^(?<first>\d+)-(?<last>\d+)$") {
                        $first = [int]$Matches["first"]
                        $last = [int]$Matches["last"]
                        if ($first -gt $last) {
                            throw "Descending range '$token'."
                        }
                        foreach ($index in $first..$last) {
                            $indexes.Add($index)
                        }
                    }
                    else {
                        $value = 0
                        if (-not [int]::TryParse($token, [ref]$value)) {
                            throw "Invalid index '$token'."
                        }
                        $indexes.Add($value)
                    }
                }

                $result = @($indexes | Sort-Object -Unique)
                $outside = @($result | Where-Object { $_ -lt $Minimum -or $_ -gt $Maximum })
                if ($outside.Count -gt 0) {
                    throw "Indexes outside $Minimum-${Maximum}: $(@($outside) -join ', ')."
                }
                return [int[]]$result
            }
            catch {
                Write-Host "$($_.Exception.Message) Enter comma-separated indexes or ranges, or press Enter for none." -ForegroundColor Yellow
            }
        }
    }

    if ((Test-Path -LiteralPath $ManifestPath) -and
        -not (Test-Path -LiteralPath $ManifestPath -PathType Leaf)) {
        throw "Totality frame manifest path is not a file: '$ManifestPath'."
    }
    $manifestExists = Test-Path -LiteralPath $ManifestPath -PathType Leaf
    $sourceManifest = $null
    if ($manifestExists) {
        try {
            $sourceManifest = Get-Content -LiteralPath $ManifestPath -Raw -Encoding UTF8 |
                ConvertFrom-Json -ErrorAction Stop
        }
        catch {
            throw "Cannot parse totality frame manifest '$ManifestPath': $($_.Exception.Message)"
        }
        if ($null -eq $sourceManifest -or $sourceManifest -isnot [PSCustomObject]) {
            throw "Totality frame manifest must contain one JSON object: '$ManifestPath'."
        }
    }
    else {
        $sourceManifest = [PSCustomObject]@{}
    }

    $propertyNames = @($sourceManifest.PSObject.Properties.Name)
    $needsWrite = -not $manifestExists

    if ($propertyNames -contains "BlockLength") {
        $blockLength = [int]$sourceManifest.BlockLength
    }
    else {
        $blockLength = Read-ManifestInteger -Prompt "Nominal block length" -Minimum 3 -Maximum 1000000
        $needsWrite = $true
    }
    if ($propertyNames -contains "BlockStartIndex") {
        $blockStartIndex = [int]$sourceManifest.BlockStartIndex
    }
    else {
        $blockStartIndex = Read-ManifestInteger -Prompt "Block sequence start index [0 or 1]" -Minimum 0 -Maximum 1
        $needsWrite = $true
    }
    if ($propertyNames -contains "C2Index") {
        $c2Index = ConvertTo-ManifestBoundaryIndex -PropertyName "C2Index" -Value $sourceManifest.C2Index
    }
    else {
        $c2Index = Read-ManifestBoundaryIndex `
            -Prompt "C2 index (start of totality; Enter = absent)" `
            -Minimum $MinimumIndex `
            -Maximum $MaximumIndex
        $needsWrite = $true
    }
    if ($propertyNames -contains "C3Index") {
        $c3Index = ConvertTo-ManifestBoundaryIndex -PropertyName "C3Index" -Value $sourceManifest.C3Index
    }
    else {
        $c3Index = Read-ManifestBoundaryIndex `
            -Prompt "C3 index (end of totality; Enter = absent)" `
            -Minimum $MinimumIndex `
            -Maximum $MaximumIndex
        $needsWrite = $true
    }
    if ($propertyNames -contains "SkipIndices") {
        $skipIndices = @($sourceManifest.SkipIndices | ForEach-Object { [int]$_ } | Sort-Object -Unique)
    }
    else {
        $skipIndices = @(Read-ManifestIndexList `
            -Prompt "Transition indexes to skip (comma-separated indexes/ranges; Enter = none)" `
            -Minimum $MinimumIndex `
            -Maximum $MaximumIndex)
        $needsWrite = $true
    }

    if ($blockLength -lt 3) {
        throw "Manifest BlockLength must be at least 3: '$ManifestPath'."
    }
    if ($blockStartIndex -notin @(0, 1)) {
        throw "Manifest BlockStartIndex must be 0 or 1: '$ManifestPath'."
    }
    if ($null -ne $c2Index -and $null -ne $c3Index -and $c2Index -gt $c3Index) {
        throw "Manifest C2Index $c2Index must not be greater than C3Index ${c3Index}: '$ManifestPath'."
    }
    $outsideSkipIndexes = @($skipIndices | Where-Object { $_ -lt $MinimumIndex -or $_ -gt $MaximumIndex })
    if ($outsideSkipIndexes.Count -gt 0) {
        throw "Manifest SkipIndices are outside raw frame range $MinimumIndex-${MaximumIndex}: $(@($outsideSkipIndexes) -join ', ')."
    }

    $resolvedManifest = [ordered]@{
        BlockLength     = $blockLength
        BlockStartIndex = $blockStartIndex
        C2Index         = $c2Index
        C3Index         = $c3Index
    }
    $knownNames = @("BlockLength", "BlockStartIndex", "C2Index", "C3Index", "SkipIndices")
    foreach ($property in $sourceManifest.PSObject.Properties) {
        if ($property.Name -notin $knownNames) {
            $resolvedManifest[$property.Name] = $property.Value
        }
    }
    $resolvedManifest["SkipIndices"] = [int[]]$skipIndices

    $wasWritten = $false
    if ($needsWrite) {
        $manifestDirectory = Split-Path -Path $ManifestPath -Parent
        if (-not (Test-Path -LiteralPath $manifestDirectory -PathType Container)) {
            throw "Manifest parent folder not found: '$manifestDirectory'."
        }
        if ($PSCmdlet.ShouldProcess($ManifestPath, "Write completed totality frame manifest")) {
            $json = $resolvedManifest | ConvertTo-Json -Depth 6
            [System.IO.File]::WriteAllText(
                $ManifestPath,
                $json,
                [System.Text.UTF8Encoding]::new($false)
            )
            $wasWritten = $true
        }
    }

    return [PSCustomObject]@{
        Manifest     = [PSCustomObject]$resolvedManifest
        ManifestPath = $ManifestPath
        WasWritten   = $wasWritten
    }
}

function Get-AsiToPixTotalitySequenceStartIndex {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ExposureDirectory
    )

    if (-not (Test-Path -LiteralPath $ExposureDirectory -PathType Container)) {
        throw "Exposure folder not found while detecting the sequence start index: '$ExposureDirectory'."
    }

    $resolvedDirectory = (Resolve-Path -LiteralPath $ExposureDirectory -ErrorAction Stop).ProviderPath
    $namePattern = [regex]::new(
        "_(?<index>\d+)(?:_c(?:_d(?:_r)?)?)?\.(?:xisf|fit|fits|fts)$",
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )
    $foundIndexOne = $false

    foreach ($file in @(Get-ChildItem -LiteralPath $resolvedDirectory -Recurse -File -ErrorAction Stop)) {
        if ($file.Name -like "Block_*") {
            continue
        }
        if ($file.DirectoryName -match "[\\/]integrated(?:[\\/]|$)") {
            continue
        }

        $match = $namePattern.Match($file.Name)
        if (-not $match.Success) {
            continue
        }

        $index = 0
        if (-not [int]::TryParse(
                $match.Groups["index"].Value,
                [System.Globalization.NumberStyles]::None,
                [System.Globalization.CultureInfo]::InvariantCulture,
                [ref]$index)) {
            continue
        }

        if ($index -eq 0) {
            return 0
        }
        if ($index -eq 1) {
            $foundIndexOne = $true
        }
    }

    if ($foundIndexOne) {
        return 1
    }

    throw "Cannot determine whether the sequence in '$resolvedDirectory' starts at index 0 or 1. " +
        "Neither index was found outside integrated output folders. Pass -SequenceStartIndex 0 or 1 explicitly."
}

function Get-AsiToPixTotalityIntegrationPlan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$Frames,

        [Parameter(Mandatory)]
        [string]$OutputDirectory,

        [ValidateRange(0, [int]::MaxValue)]
        [int]$BlockLength = 0,

        [ValidateRange(0, 1)]
        [int]$SequenceStartIndex = 0,

        [ValidateRange(2, [int]::MaxValue)]
        [int]$MinimumFrameCount = 3
    )

    if ($Frames.Count -eq 0) {
        throw "Cannot create an integration plan from an empty frame list."
    }
    if ($BlockLength -gt 0 -and $BlockLength -lt $MinimumFrameCount) {
        throw "Block length must be zero (automatic gap splitting) or at least $MinimumFrameCount."
    }

    $sortedFrames = @($Frames | Sort-Object -Property Index, Name)
    $duplicateIndexes = @($sortedFrames | Group-Object -Property Index | Where-Object { $_.Count -gt 1 })
    if ($duplicateIndexes.Count -gt 0) {
        throw "Cannot create an integration plan with duplicate frame indexes: $(@($duplicateIndexes.Name) -join ', ')."
    }

    $groups = [System.Collections.Generic.List[object]]::new()
    if ($BlockLength -gt 0) {
        $currentBin = $null
        $currentGroup = [System.Collections.Generic.List[object]]::new()
        foreach ($frame in $sortedFrames) {
            if ($frame.Index -lt $SequenceStartIndex) {
                throw "Frame index $($frame.Index) is before sequence start index $SequenceStartIndex."
            }

            $bin = [int][Math]::Floor(($frame.Index - $SequenceStartIndex) / $BlockLength)
            if ($currentGroup.Count -gt 0 -and $bin -ne $currentBin) {
                $nominalFirstIndex = $SequenceStartIndex + ($currentBin * $BlockLength)
                $groups.Add([PSCustomObject]@{
                    Frames            = $currentGroup.ToArray()
                    NominalFirstIndex = $nominalFirstIndex
                    NominalLastIndex  = $nominalFirstIndex + $BlockLength - 1
                })
                $currentGroup = [System.Collections.Generic.List[object]]::new()
            }
            $currentGroup.Add($frame)
            $currentBin = $bin
        }
        if ($currentGroup.Count -gt 0) {
            $nominalFirstIndex = $SequenceStartIndex + ($currentBin * $BlockLength)
            $groups.Add([PSCustomObject]@{
                Frames            = $currentGroup.ToArray()
                NominalFirstIndex = $nominalFirstIndex
                NominalLastIndex  = $nominalFirstIndex + $BlockLength - 1
            })
        }
    }
    else {
        $currentGroup = [System.Collections.Generic.List[object]]::new()
        $previousIndex = $null
        foreach ($frame in $sortedFrames) {
            if ($currentGroup.Count -gt 0 -and $frame.Index -ne ($previousIndex + 1)) {
                $groups.Add([PSCustomObject]@{
                    Frames            = $currentGroup.ToArray()
                    NominalFirstIndex = $null
                    NominalLastIndex  = $null
                })
                $currentGroup = [System.Collections.Generic.List[object]]::new()
            }
            $currentGroup.Add($frame)
            $previousIndex = $frame.Index
        }
        if ($currentGroup.Count -gt 0) {
            $groups.Add([PSCustomObject]@{
                Frames            = $currentGroup.ToArray()
                NominalFirstIndex = $null
                NominalLastIndex  = $null
            })
        }
    }

    $fullOutputDirectory = [System.IO.Path]::GetFullPath($OutputDirectory)
    $plan = [System.Collections.Generic.List[object]]::new()
    for ($blockIndex = 0; $blockIndex -lt $groups.Count; $blockIndex++) {
        $group = $groups[$blockIndex]
        $blockFrames = @($group.Frames)
        $firstIndex = [int]$blockFrames[0].Index
        $lastIndex = [int]$blockFrames[-1].Index
        $blockNumber = $blockIndex + 1
        $outputName = "Block_{0:D3}_{1:D5}_{2:D5}.xisf" -f $blockNumber, $firstIndex, $lastIndex
        $isContiguous = ($lastIndex - $firstIndex + 1) -eq $blockFrames.Count
        $missingIndexes = @(
            if (-not $isContiguous) {
                $presentIndexes = [System.Collections.Generic.HashSet[int]]::new()
                foreach ($frame in $blockFrames) {
                    $null = $presentIndexes.Add([int]$frame.Index)
                }
                $firstIndex..$lastIndex | Where-Object { -not $presentIndexes.Contains($_) }
            }
        )
        $skipReasons = [System.Collections.Generic.List[string]]::new()
        if ($blockFrames.Count -lt $MinimumFrameCount) {
            $skipReasons.Add("fewer than $MinimumFrameCount frames")
        }
        if (-not $isContiguous) {
            $missingText = @($missingIndexes | Select-Object -First 10) -join ", "
            if ($missingIndexes.Count -gt 10) {
                $missingText += ", ..."
            }
            $skipReasons.Add("non-contiguous frame indexes; missing inside block: $missingText")
        }
        $canIntegrate = $skipReasons.Count -eq 0

        $plan.Add([PSCustomObject]@{
            BlockNumber       = $blockNumber
            NominalFirstIndex = $group.NominalFirstIndex
            NominalLastIndex  = $group.NominalLastIndex
            FirstIndex        = $firstIndex
            LastIndex         = $lastIndex
            FrameCount        = $blockFrames.Count
            Frames            = $blockFrames
            IsContiguous      = $isContiguous
            MissingIndexes    = $missingIndexes
            OutputPath        = Join-Path -Path $fullOutputDirectory -ChildPath $outputName
            CanIntegrate      = $canIntegrate
            SkipReason        = if ($canIntegrate) { $null } else { @($skipReasons) -join "; " }
            SplitMode         = if ($BlockLength -gt 0) { "NominalIndexRanges" } else { "IndexGaps" }
        })
    }

    return @($plan)
}

function Get-AsiToPixTotalityFrameManifestPlan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [object[]]$RawFrames,

        [AllowEmptyCollection()]
        [Alias("RegisteredFrames")]
        [object[]]$InputFrames = @(),

        [Parameter(Mandatory)]
        [object]$Manifest,

        [Parameter(Mandatory)]
        [string]$OutputDirectory,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [ValidatePattern('^[^<>:"/\\|?*]+$')]
        [string]$ExposureLabel,

        [ValidateSet("Registered", "Debayered")]
        [string]$InputStage = "Registered",

        [ValidateRange(2, [int]::MaxValue)]
        [int]$MinimumFrameCount = 3
    )

    if ($RawFrames.Count -eq 0) {
        throw "Cannot create an integration plan without raw frames."
    }

    $sortedRawFrames = @($RawFrames | Sort-Object -Property Index, Name)
    $rawDuplicates = @($sortedRawFrames | Group-Object -Property Index | Where-Object { $_.Count -gt 1 })
    if ($rawDuplicates.Count -gt 0) {
        throw "Cannot create an integration plan with duplicate raw indexes: $(@($rawDuplicates.Name) -join ', ')."
    }
    $sortedInputFrames = @($InputFrames | Sort-Object -Property Index, Name)
    $inputDuplicates = @($sortedInputFrames | Group-Object -Property Index | Where-Object { $_.Count -gt 1 })
    if ($inputDuplicates.Count -gt 0) {
        throw "Cannot create an integration plan with duplicate input indexes: $(@($inputDuplicates.Name) -join ', ')."
    }

    $blockLength = [int]$Manifest.BlockLength
    $blockStartIndex = [int]$Manifest.BlockStartIndex
    $c2Index = if ($null -eq $Manifest.C2Index) { $null } else { [int]$Manifest.C2Index }
    $c3Index = if ($null -eq $Manifest.C3Index) { $null } else { [int]$Manifest.C3Index }
    $skipIndices = @($Manifest.SkipIndices | ForEach-Object { [int]$_ } | Sort-Object -Unique)
    if ($blockLength -lt 3) {
        throw "Manifest BlockLength must be at least 3."
    }
    if ($blockStartIndex -notin @(0, 1)) {
        throw "Manifest BlockStartIndex must be 0 or 1."
    }
    if ($null -ne $c2Index -and $null -ne $c3Index -and $c2Index -gt $c3Index) {
        throw "Manifest C2Index $c2Index must not be greater than C3Index $c3Index."
    }

    $minimumRawIndex = [int]$sortedRawFrames[0].Index
    $maximumRawIndex = [int]$sortedRawFrames[-1].Index
    foreach ($boundary in @($c2Index, $c3Index) | Where-Object { $null -ne $_ }) {
        if ($boundary -lt $minimumRawIndex -or $boundary -gt $maximumRawIndex) {
            throw "Manifest phase boundary index $boundary is outside raw frame range $minimumRawIndex-$maximumRawIndex."
        }
    }

    $rawIndexes = [System.Collections.Generic.HashSet[int]]::new()
    foreach ($frame in $sortedRawFrames) {
        $null = $rawIndexes.Add([int]$frame.Index)
    }
    $inputByIndex = @{}
    foreach ($frame in $sortedInputFrames) {
        $inputByIndex[[int]$frame.Index] = $frame
    }
    $skipIndexSet = [System.Collections.Generic.HashSet[int]]::new()
    foreach ($index in $skipIndices) {
        $null = $skipIndexSet.Add($index)
    }

    $firstRawBin = [int][Math]::Floor(($minimumRawIndex - $blockStartIndex) / $blockLength)
    $lastRawBin = [int][Math]::Floor(($maximumRawIndex - $blockStartIndex) / $blockLength)
    if ($null -eq $c2Index) {
        $firstIncludedBin = $firstRawBin
        $firstTotalityBin = $firstRawBin
    }
    else {
        $c2Bin = [int][Math]::Floor(($c2Index - $blockStartIndex) / $blockLength)
        $c2NominalFirst = $blockStartIndex + ($c2Bin * $blockLength)
        $firstIncludedBin = [Math]::Max($firstRawBin, $c2Bin - 1)
        $firstTotalityBin = if ($c2Index -eq $c2NominalFirst) { $c2Bin } else { $c2Bin + 1 }
    }
    if ($null -eq $c3Index) {
        $lastIncludedBin = $lastRawBin
        $lastTotalityBin = $lastRawBin
    }
    else {
        $c3Bin = [int][Math]::Floor(($c3Index - $blockStartIndex) / $blockLength)
        $c3NominalFirst = $blockStartIndex + ($c3Bin * $blockLength)
        $c3NominalLast = $c3NominalFirst + $blockLength - 1
        $lastIncludedBin = [Math]::Min($lastRawBin, $c3Bin + 1)
        $lastTotalityBin = if ($c3Index -eq $c3NominalLast) { $c3Bin } else { $c3Bin - 1 }
    }

    $fullOutputDirectory = [System.IO.Path]::GetFullPath($OutputDirectory)
    $plan = [System.Collections.Generic.List[object]]::new()
    $blockNumber = 0
    foreach ($bin in $firstIncludedBin..$lastIncludedBin) {
        $blockNumber++
        $nominalFirstIndex = $blockStartIndex + ($bin * $blockLength)
        $nominalLastIndex = $nominalFirstIndex + $blockLength - 1
        $rangeFirstIndex = [Math]::Max($minimumRawIndex, $nominalFirstIndex)
        $rangeLastIndex = [Math]::Min($maximumRawIndex, $nominalLastIndex)
        $phase = if ($bin -lt $firstTotalityBin) {
            "BeadsC2"
        }
        elseif ($bin -gt $lastTotalityBin) {
            "BeadsC3"
        }
        else {
            "Totality"
        }

        $rangeIndexes = @($rangeFirstIndex..$rangeLastIndex)
        $skippedInBlock = @($rangeIndexes | Where-Object { $skipIndexSet.Contains($_) })
        $integrateIndexes = @($rangeIndexes | Where-Object { -not $skipIndexSet.Contains($_) })
        $missingRawIndexes = @($integrateIndexes | Where-Object { -not $rawIndexes.Contains($_) })
        $missingInputIndexes = @($integrateIndexes | Where-Object {
            $rawIndexes.Contains($_) -and -not $inputByIndex.ContainsKey($_)
        })
        $frames = @($integrateIndexes | Where-Object { $inputByIndex.ContainsKey($_) } | ForEach-Object {
            $inputByIndex[$_]
        })

        $firstIndex = if ($integrateIndexes.Count -gt 0) { $integrateIndexes[0] } else { $rangeFirstIndex }
        $lastIndex = if ($integrateIndexes.Count -gt 0) { $integrateIndexes[-1] } else { $rangeLastIndex }
        $outputName = "Block_{0:D3}_{1}_{2}_{3:D5}_{4:D5}.xisf" -f (
            $blockNumber,
            $ExposureLabel,
            $phase,
            $firstIndex,
            $lastIndex
        )
        $skipReasons = [System.Collections.Generic.List[string]]::new()
        if ($integrateIndexes.Count -lt $MinimumFrameCount) {
            $skipReasons.Add("fewer than $MinimumFrameCount integration indexes")
        }
        if ($missingRawIndexes.Count -gt 0) {
            $skipReasons.Add("raw indexes missing: $(Format-AsiToPixTotalityIndexList -Indexes $missingRawIndexes)")
        }
        if ($missingInputIndexes.Count -gt 0) {
            $incompleteStage = if ($InputStage -eq "Debayered") { "debayering" } else { "registration" }
            $skipReasons.Add("$incompleteStage incomplete: $(Format-AsiToPixTotalityIndexList -Indexes $missingInputIndexes)")
        }
        $canIntegrate = $skipReasons.Count -eq 0

        $plan.Add([PSCustomObject]@{
            BlockNumber             = $blockNumber
            ExposureLabel           = $ExposureLabel
            InputStage              = $InputStage
            Phase                   = $phase
            NominalFirstIndex       = $nominalFirstIndex
            NominalLastIndex        = $nominalLastIndex
            FirstIndex              = $firstIndex
            LastIndex               = $lastIndex
            FrameCount              = $integrateIndexes.Count
            AvailableFrameCount     = $frames.Count
            Frames                  = $frames
            IntegrateIndexes        = $integrateIndexes
            SkippedIndexes          = $skippedInBlock
            MissingRawIndexes       = $missingRawIndexes
            MissingInputIndexes     = $missingInputIndexes
            MissingRegisteredIndexes = $missingInputIndexes
            OutputPath              = Join-Path -Path $fullOutputDirectory -ChildPath $outputName
            CanIntegrate            = $canIntegrate
            SkipReason              = if ($canIntegrate) { $null } else { @($skipReasons) -join "; " }
            SplitMode               = "FrameManifest"
        })
    }

    return @($plan)
}

function Resolve-AsiToPixPixInsightPath {
    [CmdletBinding()]
    param(
        [string]$Path
    )

    if (-not [string]::IsNullOrWhiteSpace($Path)) {
        if (-not (Test-Path -LiteralPath $Path -PathType Leaf)) {
            throw "PixInsight executable not found: '$Path'."
        }
        return (Resolve-Path -LiteralPath $Path -ErrorAction Stop).ProviderPath
    }

    $candidates = [System.Collections.Generic.List[string]]::new()
    $programFiles = [Environment]::GetFolderPath([Environment+SpecialFolder]::ProgramFiles)
    if (-not [string]::IsNullOrWhiteSpace($programFiles)) {
        $candidates.Add((Join-Path -Path $programFiles -ChildPath "PixInsight\bin\PixInsight.exe"))
    }

    $command = Get-Command -Name "PixInsight.exe" -CommandType Application -ErrorAction SilentlyContinue |
        Select-Object -First 1
    if ($null -ne $command) {
        $candidates.Add($command.Source)
    }

    foreach ($candidate in $candidates) {
        if (Test-Path -LiteralPath $candidate -PathType Leaf) {
            return (Resolve-Path -LiteralPath $candidate -ErrorAction Stop).ProviderPath
        }
    }

    throw "PixInsight.exe was not found. Pass its full path with -PixInsightPath."
}

function Read-AsiToPixTotalityBlockLength {
    [CmdletBinding()]
    param(
        [scriptblock]$InputReader = { param($Message) Read-Host $Message }
    )

    while ($true) {
        $answer = ([string](& $InputReader "Nominal block size (Enter = split automatically at missing indexes)")).Trim()
        if ([string]::IsNullOrEmpty($answer)) {
            return 0
        }

        $length = 0
        if ([int]::TryParse($answer, [ref]$length) -and $length -ge 3) {
            return $length
        }
        Write-Host "Press Enter for automatic gap splitting, or enter an integer of at least 3." -ForegroundColor Yellow
    }
}

function Read-AsiToPixTotalityInputStage {
    [CmdletBinding()]
    param(
        [scriptblock]$InputReader = { param($Message) Read-Host $Message }
    )

    Write-Host "Input folder:" -ForegroundColor Cyan
    Write-Host "  [1] Registered (default)"
    Write-Host "  [2] Debayered"
    while ($true) {
        $answer = ([string](& $InputReader "Folder to integrate [1]")).Trim()
        if ([string]::IsNullOrEmpty($answer) -or $answer -eq "1" -or $answer -ieq "Registered") {
            return "Registered"
        }
        if ($answer -eq "2" -or $answer -ieq "Debayered") {
            return "Debayered"
        }
        Write-Host "Enter 1/Registered or 2/Debayered." -ForegroundColor Yellow
    }
}

function Read-AsiToPixPixInsightMode {
    [CmdletBinding()]
    param(
        [ValidateRange(0, [int]::MaxValue)]
        [int]$RunningInstanceCount,

        [scriptblock]$InputReader = { param($Message) Read-Host $Message }
    )

    if ($RunningInstanceCount -eq 0) {
        Write-Host "No running PixInsight instance was found; a dedicated automation instance will be used." -ForegroundColor Yellow
        return "Dedicated"
    }

    Write-Host "PixInsight execution mode:" -ForegroundColor Cyan
    Write-Host "  [1] Reuse a running instance and leave it open (default)"
    Write-Host "  [2] Start a dedicated instance and close it when finished"
    if ($RunningInstanceCount -gt 1) {
        Write-Host "$RunningInstanceCount running PixInsight instances detected; Reuse targets the first available IPC slot." -ForegroundColor Yellow
    }
    while ($true) {
        $answer = ([string](& $InputReader "PixInsight mode [1]")).Trim()
        if ([string]::IsNullOrEmpty($answer) -or $answer -eq "1" -or $answer -ieq "Reuse") {
            return "Reuse"
        }
        if ($answer -eq "2" -or $answer -ieq "Dedicated") {
            return "Dedicated"
        }
        Write-Host "Enter 1/Reuse or 2/Dedicated." -ForegroundColor Yellow
    }
}

function Read-AsiToPixTotalityPlanConfirmation {
    [CmdletBinding()]
    param(
        [string]$Prompt = "Execute plan? [Y/n]",

        [scriptblock]$InputReader = { param($Message) Read-Host $Message }
    )

    while ($true) {
        $answer = ([string](& $InputReader $Prompt)).Trim()
        if ([string]::IsNullOrEmpty($answer) -or
            $answer -ceq "y" -or $answer -ceq "Y" -or
            $answer -ceq [string][char]0x0434 -or $answer -ceq [string][char]0x0414) {
            return $true
        }
        if ($answer -ceq "n" -or $answer -ceq "N" -or
            $answer -ceq [string][char]0x043D -or $answer -ceq [string][char]0x041D) {
            return $false
        }
        Write-Host "Please answer y/n." -ForegroundColor Yellow
    }
}

function New-AsiToPixTotalityManifest {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory)]
        [object[]]$Blocks,

        [Parameter(Mandatory)]
        [string]$ManifestDirectory
    )

    if (-not (Test-Path -LiteralPath $ManifestDirectory -PathType Container)) {
        throw "Manifest directory not found: '$ManifestDirectory'."
    }
    if ($Blocks.Count -eq 0) {
        throw "Cannot write a PixInsight integration manifest without blocks."
    }

    $manifestBlocks = [System.Collections.Generic.List[object]]::new()
    foreach ($block in @($Blocks | Sort-Object -Property BlockNumber)) {
        if (Test-Path -LiteralPath $block.OutputPath) {
            throw "Integration output already exists and will not be overwritten: '$($block.OutputPath)'."
        }

        $inputPaths = @($block.Frames | ForEach-Object {
            if (-not (Test-Path -LiteralPath $_.Path -PathType Leaf)) {
                throw "Integration input frame not found: '$($_.Path)'."
            }
            $_.Path
        })
        if ($inputPaths.Count -lt 2) {
            throw "At least two input frames are required for ImageIntegration block $($block.BlockNumber)."
        }

        $manifestBlocks.Add([ordered]@{
            blockNumber = [int]$block.BlockNumber
            outputPath  = [string]$block.OutputPath
            inputFiles  = [string[]]$inputPaths
        })
    }

    $statusPath = Join-Path -Path $ManifestDirectory -ChildPath "IntegrationStatus.json"
    $manifest = [ordered]@{
        schemaVersion = 2
        statusPath    = $statusPath
        blocks        = @($manifestBlocks)
    }
    $manifestPath = Join-Path -Path $ManifestDirectory -ChildPath "IntegrationPlan.json"
    $json = $manifest | ConvertTo-Json -Depth 6
    $utf8WithoutBom = [System.Text.UTF8Encoding]::new($false)
    if ($PSCmdlet.ShouldProcess($manifestPath, "Write PixInsight integration manifest")) {
        [System.IO.File]::WriteAllText($manifestPath, $json, $utf8WithoutBom)
    }
    return [PSCustomObject]@{
        ManifestPath = $manifestPath
        StatusPath   = $statusPath
    }
}

function Get-AsiToPixPixInsightArgumentList {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ScriptPath,

        [Parameter(Mandatory)]
        [string]$ManifestPath,

        [ValidateSet("Dedicated", "Reuse")]
        [string]$Mode = "Dedicated",

        [string]$IpcScriptPath
    )

    $pathsToValidate = @($ScriptPath, $ManifestPath)
    if ($Mode -eq "Reuse") {
        if ([string]::IsNullOrWhiteSpace($IpcScriptPath)) {
            throw "IpcScriptPath is required when reusing a running PixInsight instance."
        }
        $pathsToValidate += $IpcScriptPath
    }
    foreach ($value in $pathsToValidate) {
        if ($value.Contains(",")) {
            throw "PixInsight automation paths cannot contain commas: '$value'."
        }
        if ($value.Contains('"')) {
            throw "PixInsight automation paths cannot contain double quotes: '$value'."
        }
    }

    if ($Mode -eq "Reuse") {
        return [string[]]@(
            ('--execute="{0}"' -f $IpcScriptPath)
        )
    }

    return [string[]]@(
        "-n",
        "--automation-mode",
        ('-r="{0},manifest={1}"' -f $ScriptPath, $ManifestPath),
        "--force-exit"
    )
}

function Invoke-AsiToPixTotalityIntegrationPlan {
    [CmdletBinding(SupportsShouldProcess = $true, ConfirmImpact = "Medium")]
    param(
        [Parameter(Mandatory)]
        [object[]]$Plan,

        [Parameter(Mandatory)]
        [string]$PixInsightPath,

        [Parameter(Mandatory)]
        [string]$PixInsightScriptPath,

        [ValidateSet("Dedicated", "Reuse")]
        [string]$PixInsightMode = "Dedicated"
    )

    if (-not (Test-Path -LiteralPath $PixInsightPath -PathType Leaf)) {
        throw "PixInsight executable not found: '$PixInsightPath'."
    }
    if (-not (Test-Path -LiteralPath $PixInsightScriptPath -PathType Leaf)) {
        throw "PixInsight integration script not found: '$PixInsightScriptPath'."
    }
    if ($PixInsightMode -eq "Reuse" -and
        @(Get-Process -Name "PixInsight" -ErrorAction SilentlyContinue).Count -eq 0) {
        throw "Cannot reuse PixInsight because no running PixInsight instance was found. Start PixInsight or use -PixInsightMode Dedicated."
    }

    # Validate automation paths before creating an output directory or
    # launching PixInsight.
    $validationManifestPath = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath "AsiToPix-Totality-validation.json"
    $validationIpcScriptPath = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath "AsiToPix-Totality-validation.js"
    $null = Get-AsiToPixPixInsightArgumentList `
        -ScriptPath $PixInsightScriptPath `
        -ManifestPath $validationManifestPath `
        -Mode $PixInsightMode `
        -IpcScriptPath $validationIpcScriptPath

    $integratedCount = 0
    $existingCount = 0
    $rejectedCount = 0
    $notApprovedCount = 0
    $approvedBlocks = [System.Collections.Generic.List[object]]::new()

    foreach ($block in @($Plan | Sort-Object -Property BlockNumber)) {
        if (-not $block.CanIntegrate) {
            Write-Warning "Skipping block $($block.BlockNumber) ($($block.FirstIndex)-$($block.LastIndex)): $($block.SkipReason)."
            $rejectedCount++
            continue
        }

        if (Test-Path -LiteralPath $block.OutputPath) {
            if (-not (Test-Path -LiteralPath $block.OutputPath -PathType Leaf)) {
                throw "Integration output path is not a file: '$($block.OutputPath)'."
            }
            Write-Warning "Integration output already exists and will not be overwritten: '$($block.OutputPath)'."
            $existingCount++
            continue
        }

        foreach ($frame in $block.Frames) {
            if (-not (Test-Path -LiteralPath $frame.Path -PathType Leaf)) {
                throw "Integration input frame not found: '$($frame.Path)'."
            }
        }

        $inputStageDescription = if (@($block.PSObject.Properties.Name) -contains "InputStage") {
            ([string]$block.InputStage).ToLowerInvariant()
        }
        else {
            "processed"
        }
        $description = "Integrate $($block.FrameCount) $inputStageDescription frames with PixInsight"
        if (-not $PSCmdlet.ShouldProcess($block.OutputPath, $description)) {
            $notApprovedCount++
            continue
        }

        $approvedBlocks.Add($block)
    }

    if ($approvedBlocks.Count -eq 0) {
        return [PSCustomObject]@{
            IntegratedCount  = $integratedCount
            ExistingCount    = $existingCount
            RejectedCount    = $rejectedCount
            NotApprovedCount = $notApprovedCount
        }
    }

    foreach ($block in $approvedBlocks) {
        $outputDirectory = Split-Path -Path $block.OutputPath -Parent
        if (Test-Path -LiteralPath $outputDirectory) {
            if (-not (Test-Path -LiteralPath $outputDirectory -PathType Container)) {
                throw "Integration output directory path is not a directory: '$outputDirectory'."
            }
        }
        else {
            New-Item -ItemType Directory -Path $outputDirectory -Force -ErrorAction Stop | Out-Null
        }
    }

    $temporaryDirectory = Join-Path -Path ([System.IO.Path]::GetTempPath()) -ChildPath (
        "AsiToPix-Totality-{0}" -f [guid]::NewGuid().ToString("N")
    )
    $manifestInfo = $null
    $standardOutputPath = Join-Path -Path $temporaryDirectory -ChildPath "PixInsight.stdout.txt"
    $standardErrorPath = Join-Path -Path $temporaryDirectory -ChildPath "PixInsight.stderr.txt"
    $ipcScriptPath = Join-Path -Path $temporaryDirectory -ChildPath "IntegrateTotality-IPC.js"
    $retainDiagnostics = $false
    try {
        New-Item -ItemType Directory -Path $temporaryDirectory -ErrorAction Stop | Out-Null
        $manifestInfo = New-AsiToPixTotalityManifest `
            -Blocks $approvedBlocks.ToArray() `
            -ManifestDirectory $temporaryDirectory `
            -Confirm:$false
        if ($PixInsightMode -eq "Reuse") {
            $scriptText = [System.IO.File]::ReadAllText($PixInsightScriptPath, [System.Text.Encoding]::UTF8)
            $ipcMarker = "// ASITOPIX_IPC_MANIFEST"
            if (-not $scriptText.Contains($ipcMarker)) {
                throw "PixInsight integration script has no IPC manifest marker: '$PixInsightScriptPath'."
            }
            $manifestPathLiteral = ConvertTo-Json -InputObject ([string]$manifestInfo.ManifestPath) -Compress
            $ipcDeclaration = "var ASITOPIX_MANIFEST_PATH = $manifestPathLiteral;"
            $ipcScriptText = $scriptText.Replace($ipcMarker, $ipcDeclaration)
            [System.IO.File]::WriteAllText(
                $ipcScriptPath,
                $ipcScriptText,
                [System.Text.UTF8Encoding]::new($false)
            )
        }
        $arguments = Get-AsiToPixPixInsightArgumentList `
            -ScriptPath $PixInsightScriptPath `
            -ManifestPath $manifestInfo.ManifestPath `
            -Mode $PixInsightMode `
            -IpcScriptPath $ipcScriptPath

        if ($PixInsightMode -eq "Reuse") {
            Write-Host "Sending $($approvedBlocks.Count) integration block(s) to a running PixInsight instance; it will remain open." -ForegroundColor Cyan
        }
        else {
            Write-Host "Starting one dedicated PixInsight instance for $($approvedBlocks.Count) integration block(s)." -ForegroundColor Cyan
        }
        $process = Start-Process `
            -FilePath $PixInsightPath `
            -ArgumentList $arguments `
            -WindowStyle Hidden `
            -RedirectStandardOutput $standardOutputPath `
            -RedirectStandardError $standardErrorPath `
            -PassThru `
            -ErrorAction Stop

        $lastReportedBlockNumber = $null
        $lastProgressStatus = $null
        $ipcDeadline = [DateTime]::UtcNow.AddHours(12)
        while ($true) {
            if (Test-Path -LiteralPath $manifestInfo.StatusPath -PathType Leaf) {
                try {
                    $progressStatus = Get-Content -LiteralPath $manifestInfo.StatusPath -Raw -Encoding UTF8 |
                        ConvertFrom-Json -ErrorAction Stop
                    $lastProgressStatus = $progressStatus
                    if ($progressStatus.state -eq "running" -and
                        $null -ne $progressStatus.currentBlockNumber -and
                        $progressStatus.currentBlockNumber -ne $lastReportedBlockNumber) {
                        $currentBlock = @($approvedBlocks | Where-Object {
                            $_.BlockNumber -eq $progressStatus.currentBlockNumber
                        })[0]
                        Write-Host (
                            "Integrating block {0}: {1}-{2} ({3} frames)" -f
                            $currentBlock.BlockNumber,
                            $currentBlock.FirstIndex,
                            $currentBlock.LastIndex,
                            $currentBlock.FrameCount
                        ) -ForegroundColor Cyan
                        $lastReportedBlockNumber = $progressStatus.currentBlockNumber
                    }
                }
                catch {
                    # The status file can be observed between truncate and rewrite.
                    $progressStatus = $null
                }
            }

            if ($null -ne $lastProgressStatus -and
                $lastProgressStatus.state -in @("completed", "failed")) {
                break
            }

            $process.Refresh()
            if ($PixInsightMode -eq "Dedicated" -and $process.HasExited) {
                break
            }
            if ($PixInsightMode -eq "Reuse") {
                if ($process.HasExited -and $process.ExitCode -ne 0 -and $null -eq $lastProgressStatus) {
                    throw "PixInsight IPC client returned exit code $($process.ExitCode) before the PJSR script started."
                }
                if (@(Get-Process -Name "PixInsight" -ErrorAction SilentlyContinue).Count -eq 0) {
                    throw "The reused PixInsight instance closed before integration completed."
                }
                if ([DateTime]::UtcNow -ge $ipcDeadline) {
                    throw "Timed out after 12 hours while waiting for integration in the reused PixInsight instance."
                }
            }

            Start-Sleep -Milliseconds 500
        }
        if (-not $process.HasExited) {
            $process.WaitForExit()
        }

        if (-not (Test-Path -LiteralPath $manifestInfo.StatusPath -PathType Leaf)) {
            throw "PixInsight closed before the PJSR script wrote its status file."
        }

        try {
            $status = Get-Content -LiteralPath $manifestInfo.StatusPath -Raw -Encoding UTF8 |
                ConvertFrom-Json -ErrorAction Stop
        }
        catch {
            throw "Cannot read PixInsight status file '$($manifestInfo.StatusPath)': $($_.Exception.Message)"
        }

        if ($status.state -ne "completed") {
            $blockText = if ($null -ne $status.currentBlockNumber) {
                " Block $($status.currentBlockNumber)."
            }
            else {
                ""
            }
            $errorText = if ([string]::IsNullOrWhiteSpace([string]$status.errorMessage)) {
                "PixInsight PJSR ended in state '$($status.state)'."
            }
            else {
                [string]$status.errorMessage
            }
            throw "$errorText$blockText"
        }

        foreach ($block in $approvedBlocks) {
            if (-not (Test-Path -LiteralPath $block.OutputPath -PathType Leaf)) {
                throw "PixInsight status is completed but output was not created: '$($block.OutputPath)'."
            }
            $integratedCount++
        }

        if ($null -ne $process.ExitCode -and $process.ExitCode -ne 0) {
            Write-Warning "PixInsight returned exit code $($process.ExitCode) after completing all integration blocks."
        }
    }
    catch {
        $retainDiagnostics = $true
        throw "PixInsight integration plan failed: $($_.Exception.Message) Diagnostics retained at '$temporaryDirectory'."
    }
    finally {
        if (-not $retainDiagnostics) {
            $temporaryFiles = @(
                if ($null -ne $manifestInfo) { $manifestInfo.ManifestPath; $manifestInfo.StatusPath }
                $ipcScriptPath
                $standardOutputPath
                $standardErrorPath
            )
            foreach ($temporaryFile in $temporaryFiles) {
                if (-not [string]::IsNullOrWhiteSpace([string]$temporaryFile) -and
                    (Test-Path -LiteralPath $temporaryFile -PathType Leaf)) {
                    Remove-Item -LiteralPath $temporaryFile -Force -ErrorAction SilentlyContinue
                }
            }
            if (Test-Path -LiteralPath $temporaryDirectory -PathType Container) {
                $remainingItems = @(Get-ChildItem -LiteralPath $temporaryDirectory -Force -ErrorAction SilentlyContinue)
                if ($remainingItems.Count -eq 0) {
                    Remove-Item -LiteralPath $temporaryDirectory -Force -ErrorAction SilentlyContinue
                }
            }
        }
    }

    return [PSCustomObject]@{
        IntegratedCount  = $integratedCount
        ExistingCount    = $existingCount
        RejectedCount    = $rejectedCount
        NotApprovedCount = $notApprovedCount
    }
}

Export-ModuleMember -Function `
    Get-AsiToPixTotalityPhaseFolderName, `
    Get-AsiToPixTotalityExposureDirectory, `
    Get-AsiToPixTotalityFrame, `
    Get-AsiToPixTotalityRawFrame, `
    Format-AsiToPixTotalityIndexList, `
    Resolve-AsiToPixTotalityFrameManifest, `
    Get-AsiToPixTotalitySequenceStartIndex, `
    Get-AsiToPixTotalityIntegrationPlan, `
    Get-AsiToPixTotalityFrameManifestPlan, `
    Resolve-AsiToPixPixInsightPath, `
    Read-AsiToPixTotalityBlockLength, `
    Read-AsiToPixTotalityInputStage, `
    Read-AsiToPixPixInsightMode, `
    Read-AsiToPixTotalityPlanConfirmation, `
    New-AsiToPixTotalityManifest, `
    Get-AsiToPixPixInsightArgumentList, `
    Invoke-AsiToPixTotalityIntegrationPlan
