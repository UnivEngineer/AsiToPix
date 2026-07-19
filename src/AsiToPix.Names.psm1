Set-StrictMode -Version Latest

function ConvertTo-AsiToPixNameToken {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $normalized = $Name.ToLowerInvariant()
    $normalized = $normalized -replace '\b(nebula|galaxy|cluster|region|the)\b', ' '
    $normalized = $normalized -replace '[^a-z0-9]+', ' '

    return @($normalized.Split(' ', [System.StringSplitOptions]::RemoveEmptyEntries) | Select-Object -Unique)
}

function Get-AsiToPixCatalogIdentifierSet {
    param(
        [Parameter(Mandatory = $true)]
        [string]$Name
    )

    $identifiers = foreach ($match in [regex]::Matches(
            $Name,
            '(?i)\b(?<catalog>M|NGC|IC|SH2|SH|LBN|LDN)\s*[- ]?\s*(?<number>\d+[A-Z]?)\b')) {
        "$($match.Groups['catalog'].Value.ToUpperInvariant()) $($match.Groups['number'].Value.ToUpperInvariant())"
    }

    return @($identifiers | Sort-Object -Unique)
}

function Get-AsiToPixNameMatch {
    [CmdletBinding()]
    [OutputType([PSCustomObject[]])]
    param(
        [Parameter(Mandatory = $true)]
        [string]$DetectedName,

        [Parameter(Mandatory = $true)]
        [AllowEmptyCollection()]
        [string[]]$Candidates
    )

    if ([string]::IsNullOrWhiteSpace($DetectedName) -or $Candidates.Count -eq 0) {
        return @()
    }

    $detected = $DetectedName.Trim()
    $detectedLower = $detected.ToLowerInvariant()
    $detectedCompact = $detectedLower -replace '[^a-z0-9]+', ''
    $detectedTokens = @(ConvertTo-AsiToPixNameToken -Name $detected)
    $detectedCatalogIdentifiers = @(Get-AsiToPixCatalogIdentifierSet -Name $detected)
    $detectedCatalogKey = $detectedCatalogIdentifiers -join '|'

    $nameMatches = foreach ($candidate in $Candidates) {
        if ([string]::IsNullOrWhiteSpace($candidate)) { continue }

        $candidateLower = $candidate.ToLowerInvariant()
        $candidateCompact = $candidateLower -replace '[^a-z0-9]+', ''
        $candidateTokens = @(ConvertTo-AsiToPixNameToken -Name $candidate)
        $candidateCatalogIdentifiers = @(Get-AsiToPixCatalogIdentifierSet -Name $candidate)
        $candidateCatalogKey = $candidateCatalogIdentifiers -join '|'
        $score = 0

        if ($detectedCatalogIdentifiers.Count -gt 0) {
            if ($candidateCatalogKey -ne $detectedCatalogKey) {
                continue
            }

            $score += 500 * $detectedCatalogIdentifiers.Count
        }

        if ($candidateLower -eq $detectedLower) {
            $score += 1000
        }

        if ($candidateLower.Contains($detectedLower) -or $detectedLower.Contains($candidateLower)) {
            $score += 200
        }

        if (-not [string]::IsNullOrWhiteSpace($detectedCompact) -and
            ($candidateCompact.Contains($detectedCompact) -or $detectedCompact.Contains($candidateCompact))) {
            $score += 160
        }

        $sharedTokens = @($detectedTokens | Where-Object { $candidateTokens -contains $_ })
        if ($sharedTokens.Count -gt 0) {
            $score += 40 * $sharedTokens.Count
            if ($sharedTokens.Count -eq $detectedTokens.Count) {
                $score += 80
            }
        }

        if ($score -gt 0) {
            [PSCustomObject]@{
                Name  = $candidate
                Score = $score
            }
        }
    }

    return @($nameMatches | Sort-Object -Property @{ Expression = "Score"; Descending = $true }, Name)
}

Export-ModuleMember -Function Get-AsiToPixNameMatch
