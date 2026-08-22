function ConvertTo-AsiToPixTotalityHdrExposure {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$Text
    )

    $match = [regex]::Match(
        $Text.Trim(),
        "^(?<value>\d+(?:\.\d+)?)(?<unit>us|ms|s)$",
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )
    if (-not $match.Success) {
        throw "Invalid Totality HDR exposure '$Text'. Use a value such as 1ms, 10ms, 100ms, or 1s."
    }

    $value = [decimal]0
    if (-not [decimal]::TryParse(
            $match.Groups["value"].Value,
            [System.Globalization.NumberStyles]::AllowDecimalPoint,
            [System.Globalization.CultureInfo]::InvariantCulture,
            [ref]$value)) {
        throw "Totality HDR exposure is outside the supported decimal range: '$Text'."
    }
    if ($value -le 0) {
        throw "Totality HDR exposure must be greater than zero: '$Text'."
    }

    $unit = $match.Groups["unit"].Value.ToLowerInvariant()
    $milliseconds = switch ($unit) {
        "us" { $value / [decimal]1000 }
        "ms" { $value }
        "s" { $value * [decimal]1000 }
    }
    $key = $milliseconds.ToString(
        "0.############################",
        [System.Globalization.CultureInfo]::InvariantCulture
    )

    return [PSCustomObject]@{
        Label                = $match.Value
        Value                = $value
        Unit                 = $unit
        Milliseconds         = $milliseconds
        ExposureMilliseconds = $milliseconds
        Key                  = $key
        ExposureKey          = $key
    }
}

function Get-AsiToPixTotalityHdrFrame {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$InputPath
    )

    if (-not (Test-Path -LiteralPath $InputPath -PathType Container)) {
        throw "Totality HDR input directory not found: '$InputPath'."
    }
    $resolvedInputPath = (Resolve-Path -LiteralPath $InputPath -ErrorAction Stop).ProviderPath
    $namePattern = [regex]::new(
        "^Block_(?<block>\d+)_(?<exposure>\d+(?:\.\d+)?(?:us|ms|s))(?:_.+)?\.xisf$",
        [System.Text.RegularExpressions.RegexOptions]::IgnoreCase
    )
    $frames = [System.Collections.Generic.List[object]]::new()
    $frameByBlockAndExposure = [System.Collections.Generic.Dictionary[string, object]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )

    foreach ($file in @(Get-ChildItem -LiteralPath $resolvedInputPath -File -ErrorAction Stop)) {
        $match = $namePattern.Match($file.Name)
        if (-not $match.Success) {
            continue
        }

        $blockNumber = 0
        if (-not [int]::TryParse(
                $match.Groups["block"].Value,
                [System.Globalization.NumberStyles]::None,
                [System.Globalization.CultureInfo]::InvariantCulture,
                [ref]$blockNumber)) {
            throw "Totality HDR block number is outside the supported Int32 range: '$($file.FullName)'."
        }
        $exposure = ConvertTo-AsiToPixTotalityHdrExposure -Text $match.Groups["exposure"].Value
        $identity = "{0}|{1}" -f $blockNumber, $exposure.Key
        if ($frameByBlockAndExposure.ContainsKey($identity)) {
            $previousFrame = $frameByBlockAndExposure[$identity]
            throw "Duplicate Totality HDR frame for block $blockNumber and exposure $($exposure.Label): '$($previousFrame.Path)' and '$($file.FullName)'."
        }

        $frame = [PSCustomObject]@{
            BlockNumber         = $blockNumber
            BlockText           = $match.Groups["block"].Value
            ExposureLabel       = $exposure.Label
            ExposureKey         = $exposure.Key
            ExposureMilliseconds = $exposure.Milliseconds
            Name                = $file.Name
            Path                = $file.FullName
            Extension           = $file.Extension.TrimStart(".").ToLowerInvariant()
        }
        $frameByBlockAndExposure.Add($identity, $frame)
        $frames.Add($frame)
    }

    if ($frames.Count -eq 0) {
        throw "No Block_<number>_<exposure>_*.xisf Totality HDR frames found in '$resolvedInputPath'."
    }

    return @($frames | Sort-Object -Property BlockNumber, ExposureMilliseconds, Name)
}

function Get-AsiToPixTotalityHdrExposure {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$Frames
    )

    $exposures = foreach ($group in @($Frames | Group-Object -Property ExposureKey)) {
        $first = @($group.Group | Sort-Object -Property ExposureLabel, Name)[0]
        [PSCustomObject]@{
            Label        = [string]$first.ExposureLabel
            Key          = [string]$first.ExposureKey
            Milliseconds = [decimal]$first.ExposureMilliseconds
        }
    }
    return @($exposures | Sort-Object -Property Milliseconds, Label)
}

function Get-AsiToPixTotalityHdrLadder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$Frames
    )

    $exposures = @(Get-AsiToPixTotalityHdrExposure -Frames $Frames)
    if ($exposures.Count -lt 2) {
        return @()
    }

    $blockGroups = @($Frames | Group-Object -Property BlockNumber)
    $ladders = [System.Collections.Generic.List[object]]::new()
    for ($length = $exposures.Count; $length -ge 2; $length--) {
        for ($start = 0; $start -le ($exposures.Count - $length); $start++) {
            $ladderExposures = @($exposures[$start..($start + $length - 1)])
            $completeBlockCount = 0
            foreach ($blockGroup in $blockGroups) {
                $blockExposureKeys = @($blockGroup.Group.ExposureKey)
                $missingKeys = @($ladderExposures | Where-Object {
                    $blockExposureKeys -notcontains $_.Key
                })
                if ($missingKeys.Count -eq 0) {
                    $completeBlockCount++
                }
            }

            if ($completeBlockCount -gt 0) {
                $ladders.Add([PSCustomObject]@{
                    Exposures          = $ladderExposures
                    ExposureKeys       = [string[]]@($ladderExposures.Key)
                    Label              = @($ladderExposures.Label) -join " + "
                    CompleteBlockCount = $completeBlockCount
                })
            }
        }
    }

    return @($ladders)
}

function Read-AsiToPixTotalityHdrLadder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [object[]]$Ladders,

        [scriptblock]$InputReader = { param($Message) Read-Host $Message }
    )

    Write-Host "Available HDR exposure ladders:" -ForegroundColor Cyan
    for ($index = 0; $index -lt $Ladders.Count; $index++) {
        Write-Host (
            "  [{0}] {1} ({2} complete block(s))" -f
            ($index + 1),
            $Ladders[$index].Label,
            $Ladders[$index].CompleteBlockCount
        )
    }

    while ($true) {
        $answer = ([string](& $InputReader "HDR ladder [1]")).Trim()
        if ([string]::IsNullOrEmpty($answer)) {
            return $Ladders[0]
        }

        $number = 0
        if ([int]::TryParse($answer, [ref]$number) -and
            $number -ge 1 -and $number -le $Ladders.Count) {
            return $Ladders[$number - 1]
        }

        $match = @($Ladders | Where-Object { $_.Label -ieq $answer })
        if ($match.Count -eq 1) {
            return $match[0]
        }
        Write-Host "Enter a displayed number or exact exposure ladder." -ForegroundColor Yellow
    }
}

function Select-AsiToPixTotalityHdrLadder {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$Frames,

        [string[]]$Exposure,

        [scriptblock]$InputReader = { param($Message) Read-Host $Message }
    )

    $ladders = @(Get-AsiToPixTotalityHdrLadder -Frames $Frames)
    if ($ladders.Count -eq 0) {
        $availableExposures = @(Get-AsiToPixTotalityHdrExposure -Frames $Frames)
        $availableText = if ($availableExposures.Count -eq 0) {
            "none"
        }
        else {
            @($availableExposures.Label) -join ", "
        }
        throw "No blocks contain at least two different Totality HDR exposures. Discovered exposures: $availableText."
    }

    if ($null -eq $Exposure -or $Exposure.Count -eq 0) {
        if (@(Get-AsiToPixTotalityHdrExposure -Frames $Frames).Count -eq 2) {
            return $ladders[0]
        }
        return Read-AsiToPixTotalityHdrLadder -Ladders $ladders -InputReader $InputReader
    }

    if ($Exposure.Count -lt 2) {
        throw "At least two exposures are required for a Totality HDR ladder."
    }
    $requestedExposures = @($Exposure | ForEach-Object {
        ConvertTo-AsiToPixTotalityHdrExposure -Text $_
    } | Sort-Object -Property Milliseconds)
    $requestedKeys = @($requestedExposures.Key)
    if (@($requestedKeys | Sort-Object -Unique).Count -ne $requestedKeys.Count) {
        throw "A Totality HDR exposure ladder cannot contain duplicate exposures: $($Exposure -join ', ')."
    }

    $matchingLadders = @($ladders | Where-Object {
        $_.ExposureKeys.Count -eq $requestedKeys.Count -and
        (@($_.ExposureKeys | Where-Object { $requestedKeys -notcontains $_ }).Count -eq 0)
    })
    if ($matchingLadders.Count -eq 1) {
        return $matchingLadders[0]
    }

    $availableLadders = @($ladders.Label) -join "; "
    throw "Totality HDR exposure ladder '$($Exposure -join ' + ')' is not available. Available ladders: $availableLadders."
}

function Get-AsiToPixTotalityHdrPlan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [AllowEmptyCollection()]
        [object[]]$Frames,

        [Parameter(Mandatory)]
        [ValidateNotNull()]
        [object]$Ladder,

        [Parameter(Mandatory)]
        [ValidateNotNullOrEmpty()]
        [string]$OutputDirectory
    )

    if (-not [System.IO.Path]::IsPathRooted($OutputDirectory)) {
        $OutputDirectory = Join-Path -Path $PWD.Path -ChildPath $OutputDirectory
    }
    $fullOutputDirectory = [System.IO.Path]::GetFullPath($OutputDirectory)
    $selectedExposures = @($Ladder.Exposures | Sort-Object -Property Milliseconds -Descending)
    if ($selectedExposures.Count -lt 2) {
        throw "At least two exposures are required to build a Totality HDR plan."
    }

    $plan = [System.Collections.Generic.List[object]]::new()
    foreach ($blockGroup in @($Frames | Group-Object -Property BlockNumber | Sort-Object {
        [int]$_.Name
    })) {
        $blockNumber = [int]$blockGroup.Name
        $selectedFrames = [System.Collections.Generic.List[object]]::new()
        $missingLabels = [System.Collections.Generic.List[string]]::new()
        foreach ($exposure in $selectedExposures) {
            $match = @($blockGroup.Group | Where-Object {
                $_.ExposureKey -eq $exposure.Key
            })
            if ($match.Count -eq 1) {
                $selectedFrames.Add($match[0])
            }
            else {
                $missingLabels.Add([string]$exposure.Label)
            }
        }

        $canCompose = $missingLabels.Count -eq 0
        $outputName = "Block-{0:D3}-HDR.xisf" -f $blockNumber
        $exposureLabels = [string[]]@($selectedExposures.Label)
        $plan.Add([PSCustomObject]@{
            BlockNumber     = $blockNumber
            Frames          = $selectedFrames.ToArray()
            FrameCount      = $selectedFrames.Count
            ExposureLabels  = $exposureLabels
            MissingExposures = $missingLabels.ToArray()
            OutputPath      = Join-Path -Path $fullOutputDirectory -ChildPath $outputName
            CanCompose      = $canCompose
            CanIntegrate    = $canCompose
            SkipReason      = if ($canCompose) { $null } else { "missing exposure(s): $($missingLabels -join ', ')" }
            InputStage      = "HDR"
            FirstIndex      = $exposureLabels[0]
            LastIndex       = $exposureLabels[-1]
        })
    }

    return @($plan)
}

Export-ModuleMember -Function `
    ConvertTo-AsiToPixTotalityHdrExposure, `
    Get-AsiToPixTotalityHdrFrame, `
    Get-AsiToPixTotalityHdrExposure, `
    Get-AsiToPixTotalityHdrLadder, `
    Read-AsiToPixTotalityHdrLadder, `
    Select-AsiToPixTotalityHdrLadder, `
    Get-AsiToPixTotalityHdrPlan
