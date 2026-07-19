Set-StrictMode -Version Latest

$environmentModule = Join-Path -Path $PSScriptRoot -ChildPath "AsiToPix.Environment.psm1"
Import-Module $environmentModule

$frameFoldersModule = Join-Path -Path $PSScriptRoot -ChildPath "AsiToPix.FrameFolders.psm1"
Import-Module $frameFoldersModule -Force

function Get-AsiToPixProjectSourceFolderName {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Type
    )

    foreach ($kind in @("Light", "Bias", "Dark", "Flat", "FlatDark")) {
        if (Test-AsiToPixFrameFolderName -Name $Type -Kind $kind) {
            return (Get-AsiToPixCanonicalFrameFolderName -Kind $kind)
        }
    }

    return $Type
}

function Get-AsiToPixMetadataPropertyValue {
    param(
        [AllowNull()]
        [object]$InputObject,

        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    if ($null -eq $InputObject) {
        return $null
    }

    if ($InputObject -is [System.Collections.IDictionary] -and $InputObject.Contains($Name)) {
        return $InputObject[$Name]
    }

    $property = $InputObject.PSObject.Properties[$Name]
    if ($null -eq $property) {
        return $null
    }

    return $property.Value
}

function ConvertTo-AsiToPixMetadataNumericText {
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

function Get-AsiToPixMetadataRelativeChildPath {
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

function Get-AsiToPixMetadataSourceRoot {
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

function Resolve-AsiToPixCalibrationDestination {
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath,

        [Parameter(Mandatory = $true)]
        [string]$MasterRoot
    )

    if (-not [System.IO.Path]::IsPathRooted($SourcePath)) {
        throw "Calibration source path must be absolute: '$SourcePath'."
    }
    if (-not [System.IO.Path]::IsPathRooted($MasterRoot)) {
        throw "Calibration master root must be absolute: '$MasterRoot'."
    }

    $normalizedSourcePath = [System.IO.Path]::GetFullPath($SourcePath)
    $normalizedMasterRoot = [System.IO.Path]::GetFullPath($MasterRoot)
    $relativeMasterPath = Get-AsiToPixMetadataRelativeChildPath -RootPath $normalizedMasterRoot -ChildPath $normalizedSourcePath
    if ($null -ne $relativeMasterPath) {
        return [PSCustomObject]@{
            SourceMode       = "Master"
            DestinationFolder = $normalizedSourcePath
        }
    }

    $sourceRoot = Get-AsiToPixMetadataSourceRoot -MasterRoot $normalizedMasterRoot
    $relativeSourcePath = Get-AsiToPixMetadataRelativeChildPath -RootPath $sourceRoot -ChildPath $normalizedSourcePath
    if ($null -eq $relativeSourcePath) {
        throw "Calibration source path '$SourcePath' is outside both '$sourceRoot' and '$normalizedMasterRoot'."
    }

    $destinationFolder = if ([string]::IsNullOrWhiteSpace($relativeSourcePath)) {
        $normalizedMasterRoot
    } else {
        Join-Path -Path $normalizedMasterRoot -ChildPath $relativeSourcePath
    }

    return [PSCustomObject]@{
        SourceMode       = "Source"
        DestinationFolder = $destinationFolder
    }
}

<#
.SYNOPSIS
Builds self-contained calibration source records for project_meta.json.

.DESCRIPTION
Converts selected CreateProject calibration links into records containing explicit matching values and validated Master destinations for ExportMasters.
#>
function ConvertTo-AsiToPixCalibrationSourceMetadata {
    [CmdletBinding()]
    [OutputType([object[]])]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$PendingLink,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$CameraMetadata
    )

    $records = [System.Collections.Generic.List[object]]::new()
    $recordBySource = [System.Collections.Generic.Dictionary[string, object]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    foreach ($link in $PendingLink) {
        $type = [string](Get-AsiToPixMetadataPropertyValue -InputObject $link -Name "Type")
        if ($type -notin @("Biases", "Darks", "Flats", "FlatDarks")) {
            continue
        }

        $cameraName = [string](Get-AsiToPixMetadataPropertyValue -InputObject $link -Name "Cam")
        $sourcePath = [string](Get-AsiToPixMetadataPropertyValue -InputObject $link -Name "Src")
        $tag = [string](Get-AsiToPixMetadataPropertyValue -InputObject $link -Name "Tag")
        if ([string]::IsNullOrWhiteSpace($cameraName)) {
            throw "Calibration link '$tag' is missing its camera name."
        }
        if ([string]::IsNullOrWhiteSpace($sourcePath)) {
            throw "Calibration link '$tag' for camera '$cameraName' is missing its source path."
        }

        $matchingCameras = @($CameraMetadata | Where-Object {
            $metadataCameraName = [string](Get-AsiToPixMetadataPropertyValue -InputObject $_ -Name "Name")
            $metadataCameraName.Equals($cameraName, [System.StringComparison]::OrdinalIgnoreCase)
        })
        if ($matchingCameras.Count -ne 1) {
            throw "Calibration link '$tag' identifies camera '$cameraName', which does not match exactly one Cameras entry."
        }

        $calibrationFolders = Get-AsiToPixMetadataPropertyValue -InputObject $matchingCameras[0] -Name "CalibrationFolders"
        $masterRoot = [string](Get-AsiToPixMetadataPropertyValue -InputObject $calibrationFolders -Name $type)
        if ([string]::IsNullOrWhiteSpace($masterRoot)) {
            throw "CalibrationFolders.$type is missing for camera '$cameraName'."
        }

        $gain = ConvertTo-AsiToPixMetadataNumericText -Value ([string](Get-AsiToPixMetadataPropertyValue -InputObject $link -Name "Gain"))
        $temperature = ConvertTo-AsiToPixMetadataNumericText `
            -Value ([string](Get-AsiToPixMetadataPropertyValue -InputObject $link -Name "Temperature")) `
            -UnitPattern 'C'
        $exposure = ConvertTo-AsiToPixMetadataNumericText `
            -Value ([string](Get-AsiToPixMetadataPropertyValue -InputObject $link -Name "Exposure")) `
            -UnitPattern 's'
        $filter = [string](Get-AsiToPixMetadataPropertyValue -InputObject $link -Name "Filter")
        $session = [string](Get-AsiToPixMetadataPropertyValue -InputObject $link -Name "Session")
        $target = [string](Get-AsiToPixMetadataPropertyValue -InputObject $link -Name "Target")
        $flatSetId = [string](Get-AsiToPixMetadataPropertyValue -InputObject $link -Name "FlatSetId")
        $binning = [string](Get-AsiToPixMetadataPropertyValue -InputObject $link -Name "Binning")
        $setup = [string](Get-AsiToPixMetadataPropertyValue -InputObject $link -Name "Setup")
        $angle = [string](Get-AsiToPixMetadataPropertyValue -InputObject $link -Name "Angle")
        $lightSessions = @(
            Get-AsiToPixMetadataPropertyValue -InputObject $link -Name "LightSessions" |
                Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
                Sort-Object -Unique
        )

        if ($type -in @("Darks", "FlatDarks") -and
            ($null -eq $gain -or $null -eq $temperature -or $null -eq $exposure)) {
            throw "Calibration link '$tag' is missing Gain, Temperature, or Exposure metadata."
        }
        if ($type -eq "Biases" -and ($null -eq $gain -or $null -eq $temperature)) {
            throw "Calibration link '$tag' is missing Gain or Temperature metadata."
        }
        if ($type -eq "Flats" -and
            ($null -eq $gain -or $null -eq $temperature -or [string]::IsNullOrWhiteSpace($filter))) {
            throw "Calibration link '$tag' is missing Gain, Temperature, or Filter metadata."
        }

        $destination = Resolve-AsiToPixCalibrationDestination -SourcePath $sourcePath -MasterRoot $masterRoot
        $canonicalSourcePath = [string](Get-AsiToPixMetadataPropertyValue -InputObject $link -Name "CanonicalSourcePath")
        if ([string]::IsNullOrWhiteSpace($canonicalSourcePath)) {
            $canonicalSourcePath = [System.IO.Path]::GetFullPath($sourcePath)
        } else {
            $canonicalSourcePath = [System.IO.Path]::GetFullPath($canonicalSourcePath)
        }
        $sourceKey = "$type|$canonicalSourcePath"
        if ($recordBySource.ContainsKey($sourceKey)) {
            $existingRecord = $recordBySource[$sourceKey]
            $existingRecord.LightSessions = @(
                @($existingRecord.LightSessions) + $lightSessions |
                    Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } |
                    Sort-Object -Unique
            )
            continue
        }

        $record = [PSCustomObject][ordered]@{
            Type              = $type
            Camera            = $cameraName
            SourceMode        = $destination.SourceMode
            SourcePath        = $canonicalSourcePath
            DestinationFolder = $destination.DestinationFolder
            Gain              = $gain
            TemperatureC      = $temperature
            ExposureSeconds   = $exposure
            Filter            = if ([string]::IsNullOrWhiteSpace($filter)) { $null } else { $filter.ToUpperInvariant() }
            Session           = if ([string]::IsNullOrWhiteSpace($session)) { $null } else { $session }
            Target            = if ([string]::IsNullOrWhiteSpace($target)) { $null } else { $target }
            FlatSetId         = if ([string]::IsNullOrWhiteSpace($flatSetId)) { $null } else { $flatSetId }
            Binning           = if ([string]::IsNullOrWhiteSpace($binning)) { $null } else { $binning }
            Setup             = if ([string]::IsNullOrWhiteSpace($setup)) { $null } else { $setup }
            Angle             = if ([string]::IsNullOrWhiteSpace($angle)) { $null } else { $angle }
            LightSessions     = $lightSessions
            Tag               = $tag
        }
        $records.Add($record)
        $recordBySource.Add($sourceKey, $record)
    }

    return @($records | Sort-Object Type, Camera, DestinationFolder, Tag)
}

Export-ModuleMember -Function `
    ConvertTo-AsiToPixCalibrationSourceMetadata, `
    Get-AsiToPixProjectSourceFolderName
