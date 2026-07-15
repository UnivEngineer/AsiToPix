Set-StrictMode -Version Latest

$importSessionModule = Join-Path -Path $PSScriptRoot -ChildPath "AsiToPix.ImportSession.psm1"
Import-Module $importSessionModule -Force

function ConvertTo-AsiToPixReportFilter {
    param(
        [AllowEmptyString()]
        [string]$FilterName
    )

    switch -Regex ($FilterName.Trim()) {
        '^(|None|RGB|Color|Colour)$' { return "RGB" }
        '^(L|Lum|Luminance)$' { return "L" }
        '^(H|Ha|HAlpha|H-Alpha)$' { return "H" }
        '^(S|SII|S2)$' { return "S" }
        '^(O|OIII|O3)$' { return "O" }
        default { return $null }
    }
}

function ConvertTo-AsiToPixReportNumber {
    param(
        [Parameter(Mandatory = $true)]
        [decimal]$Value
    )

    return $Value.ToString("0.############################", [System.Globalization.CultureInfo]::InvariantCulture)
}

function Format-AsiToPixIntegrationTime {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateScript({ $_ -ge 0 })]
        [decimal]$Seconds
    )

    $hours = [decimal]::Floor($Seconds / 3600)
    $remainingSeconds = $Seconds - ($hours * 3600)
    $minutes = [decimal]::Floor($remainingSeconds / 60)
    $secondsPart = $remainingSeconds - ($minutes * 60)
    $parts = [System.Collections.Generic.List[string]]::new()

    $parts.Add("$($hours.ToString([System.Globalization.CultureInfo]::InvariantCulture))h")
    $parts.Add("$($minutes.ToString('00', [System.Globalization.CultureInfo]::InvariantCulture))m")
    if ($secondsPart -gt 0) {
        $secondsText = ConvertTo-AsiToPixReportNumber -Value $secondsPart
        $parts.Add("${secondsText}s")
    }

    return $parts -join " "
}

function Format-AsiToPixHourMinute {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateScript({ $_ -ge 0 })]
        [decimal]$Seconds
    )

    $totalMinutes = [decimal]::Round($Seconds / 60, 0, [System.MidpointRounding]::AwayFromZero)
    $hours = [decimal]::Floor($totalMinutes / 60)
    $minutes = $totalMinutes - ($hours * 60)
    $hoursText = $hours.ToString([System.Globalization.CultureInfo]::InvariantCulture)
    $minutesText = $minutes.ToString("00", [System.Globalization.CultureInfo]::InvariantCulture)

    return "${hoursText}:${minutesText}"
}

function Get-AsiToPixCharacteristicExposure {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$Frame
    )

    if ($Frame.Count -eq 0) {
        return $null
    }

    return [decimal](
        $Frame |
            Group-Object ExposureSeconds |
            Sort-Object -Property @{ Expression = "Count"; Descending = $true },
                @{ Expression = { [decimal]$_.Group[0].ExposureSeconds }; Descending = $false } |
            Select-Object -First 1 -ExpandProperty Group |
            Select-Object -First 1 -ExpandProperty ExposureSeconds
    )
}

function Format-AsiToPixExposureExpression {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$Frame,

        [Parameter(Mandatory = $true)]
        [decimal]$BaseExposure
    )

    if ($Frame.Count -eq 0) {
        return "-"
    }

    $characteristicExposure = $BaseExposure
    $terms = [System.Collections.Generic.List[string]]::new()
    $exposureGroups = @(
        $Frame |
            Group-Object ExposureSeconds |
            Sort-Object -Property @{ Expression = {
                if ([decimal]$_.Group[0].ExposureSeconds -eq $characteristicExposure) { 0 } else { 1 }
            } }, @{ Expression = { [decimal]$_.Group[0].ExposureSeconds } }
    )
    foreach ($exposureGroup in $exposureGroups) {
        $exposure = [decimal]$exposureGroup.Group[0].ExposureSeconds
        $exposureText = ConvertTo-AsiToPixReportNumber -Value $exposure
        $nightCounts = @(
            $exposureGroup.Group |
                Group-Object NightDate |
                Sort-Object Name |
                ForEach-Object { $_.Count.ToString([System.Globalization.CultureInfo]::InvariantCulture) }
        )

        $countExpression = if ($nightCounts.Count -eq 1) {
            $nightCounts[0]
        } else {
            "($($nightCounts -join '+'))"
        }
        $terms.Add("$countExpression*$exposureText")
    }

    return $terms -join "+"
}

function Get-AsiToPixImportRoot {
    $roots = foreach ($drive in Get-PSDrive -PSProvider FileSystem) {
        $astroPhotoPath = Join-Path -Path $drive.Root -ChildPath "AstroPhoto"
        $importPath = Join-Path -Path $astroPhotoPath -ChildPath "Import"
        if (Test-Path -LiteralPath $importPath -PathType Container -ErrorAction SilentlyContinue) {
            (Resolve-Path -LiteralPath $importPath -ErrorAction Stop).ProviderPath
        }
    }

    return @($roots | Sort-Object -Unique)
}

function Get-AsiToPixImportReport {
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string[]]$ImportPath
    )

    $report = foreach ($root in $ImportPath) {
        if (-not (Test-Path -LiteralPath $root -PathType Container)) {
            throw "Import folder not found: $root"
        }

        $resolvedRoot = (Resolve-Path -LiteralPath $root -ErrorAction Stop).ProviderPath
        foreach ($setupDirectory in Get-ChildItem -LiteralPath $resolvedRoot -Directory -ErrorAction Stop) {
            $lightPath = Join-Path -Path $setupDirectory.FullName -ChildPath "Light"
            if (-not (Test-Path -LiteralPath $lightPath -PathType Container)) {
                continue
            }

            foreach ($objectDirectory in Get-ChildItem -LiteralPath $lightPath -Directory -ErrorAction Stop) {
                $frames = foreach ($file in Get-ChildItem -LiteralPath $objectDirectory.FullName -File -Recurse -ErrorAction Stop) {
                    if ($file.Name -notmatch '\.fits?(\.gz)?$') {
                        continue
                    }

                    $info = Get-AsiToPixLightFileInfo -FileName $file.Name
                    if ($null -eq $info.CapturedAt -or $null -eq $info.ExposureSeconds) {
                        Write-Warning "Skipping FITS file with missing timestamp or exposure: $($file.FullName)"
                        continue
                    }

                    $reportFilter = ConvertTo-AsiToPixReportFilter -FilterName $info.FilterName
                    if ($null -eq $reportFilter) {
                        Write-Warning "Skipping FITS file with unsupported filter '$($info.FilterName)': $($file.FullName)"
                        continue
                    }

                    [PSCustomObject]@{
                        Filter          = $reportFilter
                        ExposureSeconds = [decimal]::Parse(
                            $info.ExposureSeconds,
                            [System.Globalization.NumberStyles]::Number,
                            [System.Globalization.CultureInfo]::InvariantCulture
                        )
                        NightDate       = Get-AsiToPixNightDate -CapturedAt $info.CapturedAt
                    }
                }

                $frames = @($frames)
                if ($frames.Count -eq 0) {
                    continue
                }

                $baseExposure = Get-AsiToPixCharacteristicExposure -Frame $frames
                $cells = [ordered]@{}
                $filterSeconds = [ordered]@{}
                foreach ($filter in @("RGB", "L", "H", "O", "S")) {
                    $filterFrames = @($frames | Where-Object Filter -EQ $filter)
                    $cells[$filter] = Format-AsiToPixExposureExpression -Frame $filterFrames -BaseExposure $baseExposure
                    $filterSeconds[$filter] = if ($filterFrames.Count -eq 0) {
                        [decimal]0
                    } else {
                        [decimal](($filterFrames | Measure-Object -Property ExposureSeconds -Sum).Sum)
                    }
                }

                $integrationSeconds = [decimal](($frames | Measure-Object -Property ExposureSeconds -Sum).Sum)

                [PSCustomObject]@{
                    ImportRoot        = $resolvedRoot
                    Setup             = $setupDirectory.Name
                    Object            = $objectDirectory.Name
                    Exposure          = $baseExposure
                    FrameCount        = $frames.Count
                    IntegrationSeconds = $integrationSeconds
                    RGB               = $cells["RGB"]
                    L                 = $cells["L"]
                    H                 = $cells["H"]
                    O                 = $cells["O"]
                    S                 = $cells["S"]
                    RGBSeconds        = $filterSeconds["RGB"]
                    LSeconds          = $filterSeconds["L"]
                    HSeconds          = $filterSeconds["H"]
                    OSeconds          = $filterSeconds["O"]
                    SSeconds          = $filterSeconds["S"]
                }
            }
        }
    }

    return @($report | Sort-Object ImportRoot, Setup, Object)
}

function Get-AsiToPixImportReportLine {
    [CmdletBinding(DefaultParameterSetName = "Path")]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = "Path")]
        [ValidateNotNullOrEmpty()]
        [string[]]$ImportPath,

        [Parameter(Mandatory = $true, ParameterSetName = "Report")]
        [AllowEmptyCollection()]
        [object[]]$Report
    )

    $rows = if ($PSCmdlet.ParameterSetName -eq "Path") {
        @(Get-AsiToPixImportReport -ImportPath $ImportPath)
    } else {
        @($Report)
    }
    $lines = [System.Collections.Generic.List[string]]::new()
    $setupGroups = @($rows | Group-Object ImportRoot, Setup)

    for ($setupIndex = 0; $setupIndex -lt $setupGroups.Count; $setupIndex++) {
        if ($setupIndex -gt 0) {
            $lines.Add("")
        }

        $setup = $setupGroups[$setupIndex].Group[0].Setup
        $setupCells = @($setup -split '\s+@\s+', 2)
        $lines.Add($setupCells -join "`t")
        $lines.Add(@("Object", "Expo", "RGB", "L", "H", "O", "S") -join "`t")

        foreach ($row in $setupGroups[$setupIndex].Group) {
            $exposureText = ConvertTo-AsiToPixReportNumber -Value $row.Exposure
            $lines.Add(@($row.Object, $exposureText, $row.RGB, $row.L, $row.H, $row.O, $row.S) -join "`t")
        }
    }

    return $lines.ToArray()
}

function Get-AsiToPixImportReportPrettyLine {
    [CmdletBinding(DefaultParameterSetName = "Path")]
    param(
        [Parameter(Mandatory = $true, ParameterSetName = "Path")]
        [ValidateNotNullOrEmpty()]
        [string[]]$ImportPath,

        [Parameter(Mandatory = $true, ParameterSetName = "Report")]
        [AllowEmptyCollection()]
        [object[]]$Report
    )

    $rows = if ($PSCmdlet.ParameterSetName -eq "Path") {
        @(Get-AsiToPixImportReport -ImportPath $ImportPath)
    } else {
        @($Report)
    }
    $lines = [System.Collections.Generic.List[string]]::new()
    $setupGroups = @($rows | Group-Object ImportRoot, Setup)
    $headers = @("Object", "Expo", "RGB", "L", "H", "O", "S")

    for ($setupIndex = 0; $setupIndex -lt $setupGroups.Count; $setupIndex++) {
        if ($setupIndex -gt 0) {
            $lines.Add("")
        }

        $groupRows = @($setupGroups[$setupIndex].Group)
        $setupCells = @($groupRows[0].Setup -split '\s+@\s+', 2)
        $lines.Add($setupCells -join "  ")

        $tableRows = [System.Collections.Generic.List[object[]]]::new()
        $tableRows.Add($headers)
        foreach ($row in $groupRows) {
            $exposureText = ConvertTo-AsiToPixReportNumber -Value $row.Exposure
            $tableRows.Add(@($row.Object, $exposureText, $row.RGB, $row.L, $row.H, $row.O, $row.S))
        }

        $widths = for ($column = 0; $column -lt $headers.Count; $column++) {
            ($tableRows | ForEach-Object { $_[$column].Length } | Measure-Object -Maximum).Maximum
        }

        foreach ($tableRow in $tableRows) {
            $cells = for ($column = 0; $column -lt $tableRow.Count; $column++) {
                if ($column -eq $tableRow.Count - 1) {
                    $tableRow[$column]
                } else {
                    $tableRow[$column].PadRight($widths[$column])
                }
            }

            $lines.Add($cells -join "  ")
        }
    }

    return $lines.ToArray()
}

function Get-AsiToPixIntegrationSummaryPrettyLine {
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$Report
    )

    $lines = [System.Collections.Generic.List[string]]::new()
    $setupGroups = @($Report | Group-Object ImportRoot, Setup)
    $filterNames = @("RGB", "L", "H", "O", "S")

    for ($setupIndex = 0; $setupIndex -lt $setupGroups.Count; $setupIndex++) {
        if ($setupIndex -gt 0) {
            $lines.Add("")
        }

        $groupRows = @($setupGroups[$setupIndex].Group)
        $setupCells = @($groupRows[0].Setup -split '\s+@\s+', 2)
        $lines.Add($setupCells -join "  ")

        $activeFilters = @($filterNames | Where-Object {
            $propertyName = "${_}Seconds"
            (($groupRows | Measure-Object -Property $propertyName -Sum).Sum) -gt 0
        })
        $headers = @("Object") + $activeFilters
        $tableRows = [System.Collections.Generic.List[object[]]]::new()
        $tableRows.Add($headers)

        foreach ($row in $groupRows) {
            $rowCells = [System.Collections.Generic.List[string]]::new()
            $rowCells.Add($row.Object)
            foreach ($filter in $activeFilters) {
                $propertyName = "${filter}Seconds"
                $seconds = [decimal]$row.$propertyName
                $rowCells.Add($(if ($seconds -gt 0) {
                    Format-AsiToPixHourMinute -Seconds $seconds
                } else {
                    "-"
                }))
            }
            $tableRows.Add($rowCells.ToArray())
        }

        $widths = for ($column = 0; $column -lt $headers.Count; $column++) {
            ($tableRows | ForEach-Object { $_[$column].Length } | Measure-Object -Maximum).Maximum
        }
        foreach ($tableRow in $tableRows) {
            $cells = for ($column = 0; $column -lt $tableRow.Count; $column++) {
                if ($column -eq $tableRow.Count - 1) {
                    $tableRow[$column]
                } else {
                    $tableRow[$column].PadRight($widths[$column])
                }
            }
            $lines.Add($cells -join "  ")
        }
    }

    return $lines.ToArray()
}

Export-ModuleMember -Function `
    ConvertTo-AsiToPixReportFilter, `
    Format-AsiToPixHourMinute, `
    Format-AsiToPixIntegrationTime, `
    Format-AsiToPixExposureExpression, `
    Get-AsiToPixCharacteristicExposure, `
    Get-AsiToPixImportReport, `
    Get-AsiToPixImportReportLine, `
    Get-AsiToPixImportReportPrettyLine, `
    Get-AsiToPixIntegrationSummaryPrettyLine, `
    Get-AsiToPixImportRoot
