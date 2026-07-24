function Get-AsiToPixWbppCleanupPlan {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$ProcessingRoot
    )

    if (-not (Test-Path -LiteralPath $ProcessingRoot -PathType Container)) {
        throw "Processing folder not found: '$ProcessingRoot'."
    }

    $resolvedProcessingRoot = (Resolve-Path -LiteralPath $ProcessingRoot -ErrorAction Stop).ProviderPath
    $metadataFiles = @(Get-ChildItem -LiteralPath $resolvedProcessingRoot -Filter "project_meta.json" -File -Recurse -ErrorAction Stop)
    $plans = [System.Collections.Generic.List[object]]::new()

    foreach ($metadataFile in $metadataFiles) {
        $projectRoot = $metadataFile.Directory.FullName
        # Metadata marks a project but is deliberately not read here. Project
        # folders may be renamed while older metadata still contains stale paths.
        $pixPath = Join-Path -Path $projectRoot -ChildPath "Pix"
        if (-not (Test-Path -LiteralPath $pixPath -PathType Container)) {
            continue
        }

        $pixDirectory = Get-Item -LiteralPath $pixPath -Force -ErrorAction Stop
        if ($pixDirectory.Attributes -band [System.IO.FileAttributes]::ReparsePoint) {
            Write-Warning "PixPath '$pixPath' is a reparse point; the project was skipped."
            continue
        }

        $items = [System.Collections.Generic.List[object]]::new()
        $pendingDirectories = [System.Collections.Generic.Queue[string]]::new()
        $pendingDirectories.Enqueue($pixPath)
        while ($pendingDirectories.Count -gt 0) {
            $currentDirectory = $pendingDirectories.Dequeue()
            foreach ($item in @(Get-ChildItem -LiteralPath $currentDirectory -Force -ErrorAction Stop)) {
                $items.Add($item)
                if ($item.PSIsContainer -and
                    -not ($item.Attributes -band [System.IO.FileAttributes]::ReparsePoint)) {
                    $pendingDirectories.Enqueue($item.FullName)
                }
            }
        }
        $filesToRemove = @($items | Where-Object {
            -not $_.PSIsContainer -and $_.Name -notlike "masterLight*.*"
        })
        $preservedFiles = @($items | Where-Object {
            -not $_.PSIsContainer -and $_.Name -like "masterLight*.*"
        })
        $reparsePoints = @($items | Where-Object {
            $_.Attributes -band [System.IO.FileAttributes]::ReparsePoint
        })

        # Reparse points are never cleanup targets. This prevents traversal or removal
        # of links that might refer to data outside the processing project.
        $filesToRemove = @($filesToRemove | Where-Object {
            -not ($_.Attributes -band [System.IO.FileAttributes]::ReparsePoint)
        })

        if ($filesToRemove.Count -eq 0) {
            continue
        }

        $plans.Add([PSCustomObject]@{
            ProjectRoot        = $projectRoot
            MetadataPath       = $metadataFile.FullName
            PixPath            = $pixPath
            FilesToRemove      = $filesToRemove
            PreservedFiles     = $preservedFiles
            SkippedReparsePoints = $reparsePoints
            Directories        = @($items | Where-Object {
                $_.PSIsContainer -and -not ($_.Attributes -band [System.IO.FileAttributes]::ReparsePoint)
            })
            BytesToRemove      = [long](($filesToRemove | Measure-Object -Property Length -Sum).Sum)
        })
    }

    return @($plans)
}

function Read-AsiToPixYesNo {
    [CmdletBinding()]
    param(
        [Parameter(Mandatory)]
        [string]$Prompt,

        [scriptblock]$InputReader = { param($Message) Read-Host $Message }
    )

    while ($true) {
        $answer = ([string](& $InputReader $Prompt)).Trim()
        if ($answer -ceq "y" -or $answer -ceq "Y" -or
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

function Invoke-AsiToPixWbppCleanupPlan {
    [CmdletBinding(SupportsShouldProcess = $true)]
    param(
        [Parameter(Mandatory, ValueFromPipeline)]
        [object[]]$Plan,

        [scriptblock]$ConfirmationReader = { param($Message) Read-Host $Message }
    )

    begin {
        $removedFileCount = 0
        $removedDirectoryCount = 0
        $declinedProjectCount = 0
        $reclaimedBytes = [long]0
        $whatIfBytes = [long]0
    }
    process {
        foreach ($projectPlan in $Plan) {
            Write-Host "`nProject: $($projectPlan.ProjectRoot)" -ForegroundColor Cyan
            Write-Host "  Pix folder       : $($projectPlan.PixPath)"
            Write-Host "  Files to remove  : $($projectPlan.FilesToRemove.Count)"
            Write-Host ("  Space to reclaim : {0:N2} GB" -f ($projectPlan.BytesToRemove / 1GB))
            Write-Host "  masterLight kept : $($projectPlan.PreservedFiles.Count)" -ForegroundColor Green
            if ($projectPlan.SkippedReparsePoints.Count -gt 0) {
                Write-Warning "$($projectPlan.SkippedReparsePoints.Count) reparse point(s) under '$($projectPlan.PixPath)' will not be removed."
            }

            $approved = $WhatIfPreference -or (Read-AsiToPixYesNo `
                -Prompt "Clean this project? [y/n]" `
                -InputReader $ConfirmationReader)
            if (-not $approved) {
                $declinedProjectCount++
                continue
            }

            if ($WhatIfPreference) {
                $whatIfBytes += $projectPlan.BytesToRemove
            }

            foreach ($file in $projectPlan.FilesToRemove) {
                if ($PSCmdlet.ShouldProcess($file.FullName, "Remove WBPP temporary file")) {
                    $fileLength = [long]$file.Length
                    Remove-Item -LiteralPath $file.FullName -Force -ErrorAction Stop
                    $removedFileCount++
                    $reclaimedBytes += $fileLength
                }
            }

            # Remove only empty, ordinary directories. Directories containing a
            # preserved masterLight file remain in place.
            $directories = @($projectPlan.Directories | Sort-Object { $_.FullName.Length } -Descending)
            foreach ($directory in $directories) {
                if (@(Get-ChildItem -LiteralPath $directory.FullName -Force -ErrorAction Stop).Count -eq 0 -and
                    $PSCmdlet.ShouldProcess($directory.FullName, "Remove empty WBPP directory")) {
                    Remove-Item -LiteralPath $directory.FullName -Force -ErrorAction Stop
                    $removedDirectoryCount++
                }
            }
        }
    }
    end {
        return [PSCustomObject]@{
            RemovedFileCount      = $removedFileCount
            RemovedDirectoryCount = $removedDirectoryCount
            DeclinedProjectCount  = $declinedProjectCount
            ReclaimedBytes        = $reclaimedBytes
            WhatIfBytes           = $whatIfBytes
        }
    }
}

Export-ModuleMember -Function `
    Get-AsiToPixWbppCleanupPlan, `
    Invoke-AsiToPixWbppCleanupPlan, `
    Read-AsiToPixYesNo
