Set-StrictMode -Version Latest

function Get-AsiToPixFrameFolderAlias {
    [CmdletBinding()]
    [OutputType([string[]])]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("Light", "Bias", "Dark", "Flat", "FlatDark")]
        [string]$Kind
    )

    switch ($Kind) {
        "Light" { return @("Light", "Lights") }
        "Bias" { return @("Bias", "Biases") }
        "Dark" { return @("Dark", "Darks") }
        "Flat" { return @("Flat", "Flats") }
        "FlatDark" { return @("FlatDark", "FlatDarks", "flat-dark", "flat-darks") }
    }
}

function Test-AsiToPixFrameFolderName {
    [CmdletBinding()]
    [OutputType([bool])]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Name,

        [Parameter(Mandatory = $true)]
        [ValidateSet("Light", "Bias", "Dark", "Flat", "FlatDark")]
        [string]$Kind
    )

    foreach ($alias in @(Get-AsiToPixFrameFolderAlias -Kind $Kind)) {
        if ($Name.Equals($alias, [System.StringComparison]::OrdinalIgnoreCase)) {
            return $true
        }
    }

    return $false
}

function Get-AsiToPixChildFrameFolder {
    [CmdletBinding()]
    [OutputType([System.IO.DirectoryInfo[]])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$Path,

        [Parameter(Mandatory = $true)]
        [ValidateSet("Light", "Bias", "Dark", "Flat", "FlatDark")]
        [string]$Kind
    )

    if (-not (Test-Path -LiteralPath $Path -PathType Container)) {
        return @()
    }

    return @(Get-ChildItem -LiteralPath $Path -Directory -ErrorAction Stop |
        Where-Object { Test-AsiToPixFrameFolderName -Name $_.Name -Kind $Kind } |
        Sort-Object -Property FullName)
}

function Get-AsiToPixCanonicalFrameFolderName {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("Light", "Bias", "Dark", "Flat", "FlatDark")]
        [string]$Kind
    )

    switch ($Kind) {
        "Light" { return "lights" }
        "Bias" { return "biases" }
        "Dark" { return "darks" }
        "Flat" { return "flats" }
        "FlatDark" { return "flat-darks" }
    }
}

function Get-AsiToPixFrameFolderRegex {
    [CmdletBinding()]
    [OutputType([string])]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateSet("Light", "Bias", "Dark", "Flat", "FlatDark")]
        [string]$Kind
    )

    $escapedAliases = @(Get-AsiToPixFrameFolderAlias -Kind $Kind | ForEach-Object {
        [regex]::Escape($_)
    })
    return "(?i:$($escapedAliases -join '|'))"
}

function Get-AsiToPixFlatFolderInfo {
    [CmdletBinding()]
    [OutputType([PSCustomObject])]
    param(
        [Parameter(Mandatory = $true)]
        [AllowEmptyString()]
        [string]$Name
    )

    $dateText = $null
    $filterName = $null
    if ($Name.Length -ge 8) {
        $dateText = $Name.Substring(0, 8)
        $remainder = $Name.Substring(8).TrimStart()
        if ($remainder -match '^(?:filt)?(?<filter>[^ _]+)') {
            $filterName = $Matches["filter"]
        }
    }

    $angleDegrees = $null
    if ($Name -match '(?i)(?:^|\s)(?<angle>\d{1,3})deg(?=\s|$)') {
        $angleDegrees = $Matches["angle"]
    }

    return [PSCustomObject]@{
        DateText     = $dateText
        FilterName   = $filterName
        AngleDegrees = $angleDegrees
    }
}

Export-ModuleMember -Function `
    Get-AsiToPixCanonicalFrameFolderName, `
    Get-AsiToPixChildFrameFolder, `
    Get-AsiToPixFlatFolderInfo, `
    Get-AsiToPixFrameFolderAlias, `
    Get-AsiToPixFrameFolderRegex, `
    Test-AsiToPixFrameFolderName
