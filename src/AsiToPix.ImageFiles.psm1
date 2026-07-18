Set-StrictMode -Version Latest

$script:AsiToPixRawImageExtensions = @(
    ".3fr", ".ari", ".arw", ".bay", ".bmq", ".cap", ".cine", ".cr2", ".cr3", ".crw",
    ".cs1", ".dc2", ".dcr", ".dng", ".erf", ".fff", ".gpr", ".ia", ".iiq", ".k25",
    ".kc2", ".kdc", ".mdc", ".mef", ".mos", ".mrw", ".nef", ".nrw", ".orf", ".pef",
    ".ptx", ".pxn", ".qtk", ".raf", ".raw", ".rdc", ".rw2", ".rwl", ".rwz", ".sr2",
    ".srf", ".srw", ".sti", ".x3f"
)

$script:AsiToPixStandardImageExtensions = @(
    ".bmp", ".dib",
    ".fit", ".fits", ".fts",
    ".j2c", ".j2k", ".jp2", ".jpc",
    ".jpe", ".jpeg", ".jpg",
    ".pbm", ".pgm", ".png", ".pnm", ".ppm",
    ".tif", ".tiff",
    ".xisf"
)

$script:AsiToPixCompressedFitsSuffixes = @(
    ".fit.fz", ".fits.fz", ".fts.fz",
    ".fit.gz", ".fits.gz", ".fts.gz"
)

function Get-AsiToPixRawImageExtension {
    return @($script:AsiToPixRawImageExtensions)
}

function Get-AsiToPixSupportedImageExtension {
    return @(($script:AsiToPixStandardImageExtensions + $script:AsiToPixRawImageExtensions) |
        Sort-Object -Unique)
}

function Test-AsiToPixRawImageFileName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileName
    )

    $extension = [System.IO.Path]::GetExtension($FileName).ToLowerInvariant()
    return ($script:AsiToPixRawImageExtensions -contains $extension)
}

function Test-AsiToPixSupportedImageFileName {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileName
    )

    $lowerName = [System.IO.Path]::GetFileName($FileName).ToLowerInvariant()
    foreach ($suffix in $script:AsiToPixCompressedFitsSuffixes) {
        if ($lowerName.EndsWith($suffix, [System.StringComparison]::Ordinal)) {
            return $true
        }
    }

    $extension = [System.IO.Path]::GetExtension($lowerName)
    return ($script:AsiToPixStandardImageExtensions -contains $extension -or
        $script:AsiToPixRawImageExtensions -contains $extension)
}

function Get-AsiToPixImageFileStem {
    param(
        [Parameter(Mandatory = $true)]
        [string]$FileName
    )

    $leafName = [System.IO.Path]::GetFileName($FileName)
    $lowerName = $leafName.ToLowerInvariant()
    $allSuffixes = @(
        $script:AsiToPixCompressedFitsSuffixes +
        $script:AsiToPixStandardImageExtensions +
        $script:AsiToPixRawImageExtensions
    )
    $suffixes = @($allSuffixes | Sort-Object -Property Length -Descending)
    foreach ($suffix in $suffixes) {
        if ($lowerName.EndsWith($suffix, [System.StringComparison]::Ordinal)) {
            return $leafName.Substring(0, $leafName.Length - $suffix.Length)
        }
    }

    return [System.IO.Path]::GetFileNameWithoutExtension($leafName)
}

Export-ModuleMember -Function `
    Get-AsiToPixImageFileStem, `
    Get-AsiToPixRawImageExtension, `
    Get-AsiToPixSupportedImageExtension, `
    Test-AsiToPixRawImageFileName, `
    Test-AsiToPixSupportedImageFileName
