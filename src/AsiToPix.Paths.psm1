function Get-AsiToPixAstroRootCandidate {
    [CmdletBinding()]
    param(
        [object[]]$FileSystemDrive
    )

    if (-not $PSBoundParameters.ContainsKey("FileSystemDrive")) {
        $FileSystemDrive = @(Get-PSDrive -PSProvider FileSystem)
    }

    $candidates = foreach ($drive in $FileSystemDrive) {
        foreach ($rootFolderName in @("AstroPhoto", "Astro")) {
            $candidatePath = Join-Path -Path $drive.Root -ChildPath $rootFolderName
            if (Test-Path -LiteralPath $candidatePath -PathType Container -ErrorAction SilentlyContinue) {
                $candidatePath
            }
        }
    }

    return @($candidates | Sort-Object -Unique)
}

function Get-AsiToPixAstroRootChildCandidate {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory = $true)]
        [ValidateNotNullOrEmpty()]
        [string]$ChildPath,

        [object[]]$FileSystemDrive
    )

    $rootParameters = @{}
    if ($PSBoundParameters.ContainsKey("FileSystemDrive")) {
        $rootParameters.FileSystemDrive = $FileSystemDrive
    }

    $candidates = foreach ($rootPath in Get-AsiToPixAstroRootCandidate @rootParameters) {
        $candidatePath = Join-Path -Path $rootPath -ChildPath $ChildPath
        if (Test-Path -LiteralPath $candidatePath -PathType Container -ErrorAction SilentlyContinue) {
            (Resolve-Path -LiteralPath $candidatePath -ErrorAction Stop).ProviderPath
        }
    }

    return @($candidates | Sort-Object -Unique)
}

function Resolve-AstroPhotoRoot {
    [CmdletBinding()]
    param(
        [string]$Purpose = "",

        [string]$SelectionPrompt = "",

        [switch]$AlwaysPrompt
    )

    $candidates = @(Get-AsiToPixAstroRootCandidate)
    $purposeSuffix = if ([string]::IsNullOrWhiteSpace($Purpose)) {
        ""
    } else {
        " for $Purpose"
    }

    if ($candidates.Count -eq 1 -and -not $AlwaysPrompt) {
        Write-Host "[INFO] Astrophotography root found${purposeSuffix}: $($candidates[0])" -ForegroundColor Cyan
        return (Resolve-Path -LiteralPath $candidates[0]).ProviderPath
    }

    if ($candidates.Count -gt 0) {
        $availableRootsPrompt = if ([string]::IsNullOrWhiteSpace($SelectionPrompt)) {
            "Available astrophotography roots${purposeSuffix}"
        } else {
            $SelectionPrompt.Trim().TrimEnd(":")
        }
        Write-Host "`n${availableRootsPrompt}:" -ForegroundColor Cyan
        for ($i = 0; $i -lt $candidates.Count; $i++) {
            Write-Host " [$i] $($candidates[$i])" -ForegroundColor White
        }

        do {
            $rootEntryPrompt = if ([string]::IsNullOrWhiteSpace($SelectionPrompt)) {
                "Select root index or enter a full root path${purposeSuffix}"
            } else {
                "Enter root index or a full root path"
            }
            $answer = (Read-Host $rootEntryPrompt).Trim('"')
            $selectedIndex = -1
            $validSelection = [int]::TryParse($answer, [ref]$selectedIndex) -and
                $selectedIndex -ge 0 -and
                $selectedIndex -lt $candidates.Count

            if ($validSelection) {
                $selectedPath = $candidates[$selectedIndex]
            } elseif (Test-Path -LiteralPath $answer -PathType Container) {
                $selectedPath = $answer
                $validSelection = $true
            } else {
                Write-Host "[!] Invalid selection. Enter a number from 0 to $($candidates.Count - 1), or an existing root path." -ForegroundColor Red
            }
        } while (-not $validSelection)

        $resolvedPath = (Resolve-Path -LiteralPath $selectedPath -ErrorAction Stop).ProviderPath
        Write-Host "[INFO] Using astrophotography root${purposeSuffix}: $resolvedPath" -ForegroundColor Cyan
        return $resolvedPath
    }

    Write-Host "`n[!] No astrophotography root found${purposeSuffix} by patterns *:\AstroPhoto and *:\Astro." -ForegroundColor Yellow
    do {
        $manualPathPrompt = if ([string]::IsNullOrWhiteSpace($SelectionPrompt)) {
            "Enter astrophotography root path manually${purposeSuffix}"
        } else {
            "$($SelectionPrompt.Trim().TrimEnd(':')) - enter a full root path"
        }
        $manualPath = (Read-Host $manualPathPrompt).Trim('"')
        if ([string]::IsNullOrWhiteSpace($manualPath)) {
            Write-Host "[!] Path cannot be empty." -ForegroundColor Red
            continue
        }

        if (Test-Path -LiteralPath $manualPath -PathType Container) {
            $resolvedPath = (Resolve-Path -LiteralPath $manualPath).ProviderPath
            Write-Host "[INFO] Using astrophotography root${purposeSuffix}: $resolvedPath" -ForegroundColor Cyan
            return $resolvedPath
        }

        Write-Host "[!] Astrophotography root not found: $manualPath" -ForegroundColor Red
    } while ($true)
}
