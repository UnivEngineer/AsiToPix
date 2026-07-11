function Resolve-AstroPhotoRoot {
    $candidates = @(Get-PSDrive -PSProvider FileSystem |
        ForEach-Object {
            Join-Path -Path $_.Root -ChildPath "AstroPhoto"
        } |
        Where-Object {
            Test-Path -LiteralPath $_ -PathType Container
        } |
        Sort-Object -Unique)

    if ($candidates.Count -eq 1) {
        Write-Host "[INFO] AstroPhoto root found: $($candidates[0])" -ForegroundColor Cyan
        return (Resolve-Path -LiteralPath $candidates[0]).ProviderPath
    }

    if ($candidates.Count -gt 1) {
        Write-Host "`nMultiple AstroPhoto roots found:" -ForegroundColor Cyan
        for ($i = 0; $i -lt $candidates.Count; $i++) {
            Write-Host " [$i] $($candidates[$i])" -ForegroundColor White
        }

        do {
            $answer = Read-Host "Select AstroPhoto root index"
            $selectedIndex = -1
            $validSelection = [int]::TryParse($answer, [ref]$selectedIndex) -and
                $selectedIndex -ge 0 -and
                $selectedIndex -lt $candidates.Count

            if (-not $validSelection) {
                Write-Host "[!] Invalid selection. Enter a number from 0 to $($candidates.Count - 1)." -ForegroundColor Red
            }
        } while (-not $validSelection)

        Write-Host "[INFO] Using AstroPhoto root: $($candidates[$selectedIndex])" -ForegroundColor Cyan
        return (Resolve-Path -LiteralPath $candidates[$selectedIndex]).ProviderPath
    }

    Write-Host "`n[!] No AstroPhoto root found by pattern *:\AstroPhoto." -ForegroundColor Yellow
    do {
        $manualPath = (Read-Host "Enter AstroPhoto root path manually").Trim('"')
        if ([string]::IsNullOrWhiteSpace($manualPath)) {
            Write-Host "[!] Path cannot be empty." -ForegroundColor Red
            continue
        }

        if (Test-Path -LiteralPath $manualPath -PathType Container) {
            $resolvedPath = (Resolve-Path -LiteralPath $manualPath).ProviderPath
            Write-Host "[INFO] Using AstroPhoto root: $resolvedPath" -ForegroundColor Cyan
            return $resolvedPath
        }

        Write-Host "[!] AstroPhoto root not found: $manualPath" -ForegroundColor Red
    } while ($true)
}
