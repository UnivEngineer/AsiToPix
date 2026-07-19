Set-StrictMode -Version Latest

function Get-AsiToPixProjectPlanPropertyValue {
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

function Resolve-AsiToPixCanonicalSourcePath {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path
    )

    $currentPath = (Resolve-Path -LiteralPath $Path -ErrorAction Stop).ProviderPath
    $visitedPaths = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )

    for ($depth = 0; $depth -lt 16; $depth++) {
        $normalizedCurrentPath = [System.IO.Path]::GetFullPath($currentPath)
        if (-not $visitedPaths.Add($normalizedCurrentPath)) {
            throw "Symbolic-link cycle detected while resolving calibration source: '$Path'."
        }

        $item = Get-Item -LiteralPath $normalizedCurrentPath -Force -ErrorAction Stop
        if ([string]::IsNullOrWhiteSpace([string]$item.LinkType)) {
            $currentPath = $item.FullName
            break
        }

        $targets = @($item.Target | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) })
        if ($targets.Count -ne 1) {
            throw "Calibration source link must have exactly one target: '$($item.FullName)'."
        }

        $targetPath = [string]$targets[0]
        if (-not [System.IO.Path]::IsPathRooted($targetPath)) {
            $targetPath = Join-Path -Path (Split-Path -Path $item.FullName -Parent) -ChildPath $targetPath
        }
        $currentPath = (Resolve-Path -LiteralPath $targetPath -ErrorAction Stop).ProviderPath

        if ($depth -eq 15) {
            throw "Too many symbolic-link levels while resolving calibration source: '$Path'."
        }
    }

    $fullPath = [System.IO.Path]::GetFullPath($currentPath)
    $pathRoot = [System.IO.Path]::GetPathRoot($fullPath)
    if ($fullPath.Equals($pathRoot, [System.StringComparison]::OrdinalIgnoreCase)) {
        return $fullPath
    }

    return $fullPath.TrimEnd([char[]]@('\', '/'))
}

function Get-AsiToPixFlatSetId {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$SourcePath
    )

    $canonicalPath = Resolve-AsiToPixCanonicalSourcePath -Path $SourcePath
    $identityText = $canonicalPath.ToUpperInvariant()
    $sha256 = [System.Security.Cryptography.SHA256]::Create()
    try {
        $bytes = [System.Text.Encoding]::UTF8.GetBytes($identityText)
        $hash = $sha256.ComputeHash($bytes)
    } finally {
        $sha256.Dispose()
    }

    return (($hash | ForEach-Object { $_.ToString('x2') }) -join '').Substring(0, 12)
}

function ConvertTo-AsiToPixFlatTagToken {
    param(
        [AllowNull()]
        [AllowEmptyString()]
        [string]$Value,

        [string]$Fallback = "Unknown"
    )

    if ([string]::IsNullOrWhiteSpace($Value)) {
        return $Fallback
    }

    $trimmedValue = $Value.Trim()
    $hasLeadingNumericMinus = $trimmedValue -match '^-\d'
    $token = $trimmedValue -replace '[\s@]+', '-'
    $token = $token -replace '[^A-Za-z0-9.\-]+', '-'
    $token = $token.Trim('-')
    if ($hasLeadingNumericMinus) {
        $token = "-$token"
    }
    if ([string]::IsNullOrWhiteSpace($token)) {
        return $Fallback
    }

    return $token
}

function ConvertTo-AsiToPixFlatSetTag {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$FlatSetId,

        [AllowEmptyString()]
        [string]$FlatDate = "",

        [AllowEmptyString()]
        [string]$Angle = "",

        [Parameter(Mandatory = $true)]
        [string]$Setup,

        [AllowEmptyString()]
        [string]$Binning = "",

        [Parameter(Mandatory = $true)]
        [string]$Filter,

        [Parameter(Mandatory = $true)]
        [string]$Target,

        [Parameter(Mandatory = $true)]
        [string]$Gain,

        [Parameter(Mandatory = $true)]
        [string]$Temperature,

        [Parameter(Mandatory = $true)]
        [string]$Camera
    )

    $parts = @(
        "FlatSet_$(ConvertTo-AsiToPixFlatTagToken -Value $FlatSetId)",
        "FlatDate_$(ConvertTo-AsiToPixFlatTagToken -Value $FlatDate)",
        "Angle_$(ConvertTo-AsiToPixFlatTagToken -Value $Angle)",
        "Setup_$(ConvertTo-AsiToPixFlatTagToken -Value $Setup)",
        "Bin_$(ConvertTo-AsiToPixFlatTagToken -Value $Binning)",
        "Filter_$(ConvertTo-AsiToPixFlatTagToken -Value $Filter)",
        "Target_$(ConvertTo-AsiToPixFlatTagToken -Value $Target)",
        "Gain_$(ConvertTo-AsiToPixFlatTagToken -Value $Gain)",
        "Temp_$(ConvertTo-AsiToPixFlatTagToken -Value $Temperature)",
        "Cam_$(ConvertTo-AsiToPixFlatTagToken -Value $Camera)"
    )

    return ($parts -join '_')
}

function Get-AsiToPixUniqueFlatPlan {
    [CmdletBinding()]
    [OutputType([object])]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [object[]]$PendingLink
    )

    $flatGroups = [System.Collections.Generic.Dictionary[string, object]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    foreach ($link in $PendingLink) {
        if ([string](Get-AsiToPixProjectPlanPropertyValue -InputObject $link -Name 'Type') -ne 'Flats') {
            continue
        }

        $sourcePath = [string](Get-AsiToPixProjectPlanPropertyValue -InputObject $link -Name 'Src')
        $canonicalPath = Resolve-AsiToPixCanonicalSourcePath -Path $sourcePath
        $link | Add-Member -NotePropertyName CanonicalSourcePath -NotePropertyValue $canonicalPath -Force

        if (-not $flatGroups.ContainsKey($canonicalPath)) {
            $flatGroups.Add($canonicalPath, [System.Collections.Generic.List[object]]::new())
        }
        $flatGroups[$canonicalPath].Add($link)
    }

    $duplicateGroups = [System.Collections.Generic.List[object]]::new()
    foreach ($canonicalPath in $flatGroups.Keys) {
        $group = @($flatGroups[$canonicalPath])
        $tags = @($group | Select-Object -ExpandProperty Tag -Unique)
        if ($tags.Count -ne 1) {
            throw "One physical flat set was assigned incompatible WBPP tags. Source: '$canonicalPath'; Tags: $($tags -join '; ')."
        }

        $sessions = @($group | ForEach-Object {
            $configuredSessions = @(Get-AsiToPixProjectPlanPropertyValue -InputObject $_ -Name 'LightSessions')
            if ($configuredSessions.Count -gt 0) {
                $configuredSessions
            } else {
                Get-AsiToPixProjectPlanPropertyValue -InputObject $_ -Name 'Session'
            }
        } | Where-Object { -not [string]::IsNullOrWhiteSpace([string]$_) } | Sort-Object -Unique)
        $group[0] | Add-Member -NotePropertyName LightSessions -NotePropertyValue $sessions -Force

        if ($group.Count -gt 1) {
            $duplicateGroups.Add([PSCustomObject]@{
                CanonicalSourcePath = $canonicalPath
                Count               = $group.Count
                Tag                 = $tags[0]
                LightSessions       = $sessions
            })
        }
    }

    $seenFlatPaths = [System.Collections.Generic.HashSet[string]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    $uniqueLinks = [System.Collections.Generic.List[object]]::new()
    foreach ($link in $PendingLink) {
        if ([string](Get-AsiToPixProjectPlanPropertyValue -InputObject $link -Name 'Type') -ne 'Flats') {
            $uniqueLinks.Add($link)
            continue
        }

        $canonicalPath = [string](Get-AsiToPixProjectPlanPropertyValue -InputObject $link -Name 'CanonicalSourcePath')
        if ($seenFlatPaths.Add($canonicalPath)) {
            $uniqueLinks.Add($flatGroups[$canonicalPath][0])
        }
    }

    $compatibilityGroups = [System.Collections.Generic.Dictionary[string, object]]::new(
        [System.StringComparer]::OrdinalIgnoreCase
    )
    foreach ($link in @($uniqueLinks | Where-Object { $_.Type -eq 'Flats' })) {
        $compatibilityKey = [string](Get-AsiToPixProjectPlanPropertyValue -InputObject $link -Name 'FlatCompatibilityKey')
        if ([string]::IsNullOrWhiteSpace($compatibilityKey)) {
            continue
        }
        if (-not $compatibilityGroups.ContainsKey($compatibilityKey)) {
            $compatibilityGroups.Add($compatibilityKey, [System.Collections.Generic.List[object]]::new())
        }
        $compatibilityGroups[$compatibilityKey].Add($link)
    }

    $separateSetGroups = [System.Collections.Generic.List[object]]::new()
    foreach ($compatibilityKey in $compatibilityGroups.Keys) {
        $group = @($compatibilityGroups[$compatibilityKey])
        if ($group.Count -gt 1) {
            $separateSetGroups.Add([PSCustomObject]@{
                CompatibilityKey = $compatibilityKey
                FlatSets         = @($group | Select-Object FlatSetId, CanonicalSourcePath, Tag)
            })
        }
    }

    return [PSCustomObject]@{
        PendingLinks      = @($uniqueLinks)
        DuplicateGroups   = @($duplicateGroups)
        SeparateSetGroups = @($separateSetGroups)
    }
}

Export-ModuleMember -Function `
    Get-AsiToPixFlatSetId, `
    Get-AsiToPixUniqueFlatPlan, `
    ConvertTo-AsiToPixFlatSetTag, `
    Resolve-AsiToPixCanonicalSourcePath
