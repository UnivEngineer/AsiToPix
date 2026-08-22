$modulePath = Join-Path -Path $PSScriptRoot -ChildPath "..\src\AsiToPix.TotalityNormalization.psm1"
$entryPointPath = Join-Path -Path $PSScriptRoot -ChildPath "..\NormalizeTotality.ps1"
$pixInsightScriptPath = Join-Path -Path $PSScriptRoot -ChildPath "..\pixinsight\NormalizeTotality.js"
Import-Module $modulePath -Force

function Add-TestTotalityComposedFrame {
    param(
        [Parameter(Mandatory)]
        [string]$Directory,

        [Parameter(Mandatory)]
        [int]$Block,

        [int]$Width = 3
    )

    $path = Join-Path -Path $Directory -ChildPath (
        "Block-{0}-HDR.xisf" -f $Block.ToString("D$Width")
    )
    Set-Content -LiteralPath $path -Value "test"
    return $path
}

Describe "Totality normalization frame discovery" {
    It "discovers composed HDR blocks in numeric order and ignores generated outputs" {
        $inputPath = Join-Path -Path $TestDrive -ChildPath "normalization-discovery"
        New-Item -ItemType Directory -Path $inputPath -Force | Out-Null
        Add-TestTotalityComposedFrame -Directory $inputPath -Block 10 | Out-Null
        Add-TestTotalityComposedFrame -Directory $inputPath -Block 2 | Out-Null
        Set-Content -LiteralPath (Join-Path $inputPath "Block-002-HDR_norm.xisf") -Value "ignore"
        Set-Content -LiteralPath (Join-Path $inputPath "notes.txt") -Value "ignore"

        $frames = @(Get-AsiToPixTotalityNormalizationFrame -InputPath $inputPath)

        $frames.Count | Should Be 2
        $frames[0].BlockNumber | Should Be 2
        $frames[0].Name | Should Be "Block-002-HDR.xisf"
        $frames[1].BlockNumber | Should Be 10
    }

    It "rejects duplicate numeric block identifiers" {
        $inputPath = Join-Path -Path $TestDrive -ChildPath "normalization-duplicates"
        New-Item -ItemType Directory -Path $inputPath -Force | Out-Null
        Add-TestTotalityComposedFrame -Directory $inputPath -Block 1 -Width 3 | Out-Null
        Add-TestTotalityComposedFrame -Directory $inputPath -Block 1 -Width 4 | Out-Null

        $caughtError = $null
        try {
            $null = @(Get-AsiToPixTotalityNormalizationFrame -InputPath $inputPath)
        }
        catch {
            $caughtError = $_
        }

        $caughtError | Should Not BeNullOrEmpty
        $caughtError.Exception.Message | Should Match "Duplicate Totality normalization frame for block 1"
    }
}

Describe "Totality normalization plan and manifest" {
    It "creates the requested _norm and _norm_cc output names" {
        $inputPath = Join-Path -Path $TestDrive -ChildPath "normalization-plan-input"
        $outputPath = Join-Path -Path $TestDrive -ChildPath "normalization-plan-output"
        New-Item -ItemType Directory -Path $inputPath -Force | Out-Null
        Add-TestTotalityComposedFrame -Directory $inputPath -Block 7 | Out-Null
        $frames = @(Get-AsiToPixTotalityNormalizationFrame -InputPath $inputPath)

        $plan = @(Get-AsiToPixTotalityNormalizationPlan `
            -Frames $frames `
            -OutputDirectory $outputPath)

        $plan.Count | Should Be 1
        $plan[0].NormalizedPath | Should Be (Join-Path $outputPath "Block-007-HDR_norm.xisf")
        $plan[0].ColorCorrectedPath | Should Be (Join-Path $outputPath "Block-007-HDR_norm_cc.xisf")
    }

    It "writes a UTF-8 manifest with every input and output path" {
        $inputPath = Join-Path -Path $TestDrive -ChildPath "normalization-manifest-input"
        $outputPath = Join-Path -Path $TestDrive -ChildPath "normalization-manifest-output"
        $manifestPath = Join-Path -Path $TestDrive -ChildPath "normalization-manifest"
        New-Item -ItemType Directory -Path $inputPath, $outputPath, $manifestPath -Force | Out-Null
        Add-TestTotalityComposedFrame -Directory $inputPath -Block 1 | Out-Null
        Add-TestTotalityComposedFrame -Directory $inputPath -Block 2 | Out-Null
        $frames = @(Get-AsiToPixTotalityNormalizationFrame -InputPath $inputPath)
        $plan = @(Get-AsiToPixTotalityNormalizationPlan `
            -Frames $frames `
            -OutputDirectory $outputPath)
        $metricsPath = Join-Path -Path $outputPath -ChildPath "normalization_metrics.csv"

        $manifestInfo = New-AsiToPixTotalityNormalizationManifest `
            -Frames $plan `
            -AllowedReferenceFiles @($frames.Path) `
            -MetricsPath $metricsPath `
            -ManifestDirectory $manifestPath `
            -Confirm:$false
        $manifest = Get-Content -LiteralPath $manifestInfo.ManifestPath -Raw -Encoding UTF8 |
            ConvertFrom-Json

        $manifest.schemaVersion | Should Be 1
        $manifest.metricsPath | Should Be $metricsPath
        @($manifest.allowedReferenceFiles).Count | Should Be 2
        @($manifest.frames).Count | Should Be 2
        $manifest.frames[0].inputPath | Should Be $frames[0].Path
        $manifest.frames[0].normalizedPath | Should Be $plan[0].NormalizedPath
        $manifest.frames[0].colorCorrectedPath | Should Be $plan[0].ColorCorrectedPath
        [System.IO.File]::ReadAllBytes($manifestInfo.ManifestPath)[0] | Should Not Be 239
    }

    It "does not create Calibrated or launch PixInsight in WhatIf mode" {
        $astroPhotoRoot = Join-Path -Path $TestDrive -ChildPath "AstroPhoto"
        $hdrPath = Join-Path -Path $astroPhotoRoot -ChildPath "SharpCap\Totality-2026\HDR"
        $inputPath = Join-Path -Path $hdrPath -ChildPath "Composed"
        $outputPath = Join-Path -Path $hdrPath -ChildPath "Calibrated"
        New-Item -ItemType Directory -Path $inputPath -Force | Out-Null
        Add-TestTotalityComposedFrame -Directory $inputPath -Block 1 | Out-Null
        $fakePixInsightPath = Join-Path -Path $TestDrive -ChildPath "PixInsight.exe"
        Set-Content -LiteralPath $fakePixInsightPath -Value "not launched"

        $null = & $entryPointPath `
            -AstroPhotoRoot $astroPhotoRoot `
            -PixInsightPath $fakePixInsightPath `
            -WhatIf `
            -Confirm:$false

        Test-Path -LiteralPath $outputPath | Should Be $false
    }

    It "hands one manifest to the running PixInsight instance and verifies every result" {
        $normalizationRoot = Join-Path -Path $TestDrive -ChildPath "normalization-ipc"
        New-Item -ItemType Directory -Path $normalizationRoot -Force | Out-Null
        $env:ASITOPIX_TOTALITY_NORMALIZATION_TEST_ROOT = $normalizationRoot

        InModuleScope AsiToPix.TotalityNormalization {
            $inputPath = Join-Path -Path $env:ASITOPIX_TOTALITY_NORMALIZATION_TEST_ROOT -ChildPath "Composed"
            $outputPath = Join-Path -Path $env:ASITOPIX_TOTALITY_NORMALIZATION_TEST_ROOT -ChildPath "Calibrated"
            New-Item -ItemType Directory -Path $inputPath -Force | Out-Null
            $inputFrame = Join-Path -Path $inputPath -ChildPath "Block-001-HDR.xisf"
            Set-Content -LiteralPath $inputFrame -Value "frame"
            $frames = @(Get-AsiToPixTotalityNormalizationFrame -InputPath $inputPath)
            $plan = @(Get-AsiToPixTotalityNormalizationPlan `
                -Frames $frames `
                -OutputDirectory $outputPath)
            $metricsPath = Join-Path -Path $outputPath -ChildPath "normalization_metrics.csv"
            New-Item -ItemType Directory -Path $outputPath -Force | Out-Null
            Set-Content -LiteralPath $metricsPath -Value "old metrics"
            Set-Content -LiteralPath $plan[0].NormalizedPath -Value "old normalized frame"
            Set-Content -LiteralPath $plan[0].ColorCorrectedPath -Value "old color-corrected frame"
            $fakePixInsightPath = Join-Path -Path $env:ASITOPIX_TOTALITY_NORMALIZATION_TEST_ROOT -ChildPath "PixInsight.exe"
            $scriptPath = Join-Path -Path $env:ASITOPIX_TOTALITY_NORMALIZATION_TEST_ROOT -ChildPath "NormalizeTotality.js"
            $sourcePixInsightScriptPath = Join-Path -Path $PSScriptRoot -ChildPath "..\pixinsight\NormalizeTotality.js"
            Set-Content -LiteralPath $fakePixInsightPath -Value "not launched"
            Copy-Item -LiteralPath $sourcePixInsightScriptPath -Destination $scriptPath

            $script:normalizationIpcArguments = @()
            $script:normalizationIpcStarted = $false
            $script:normalizationStatusPath = $null
            $script:normalizationCompletedStatus = $null
            $script:normalizationTargetPidChecks = 0
            Mock Get-Process {
                param($Name, $Id)
                $null = $Name
                if ($null -ne $Id) {
                    $script:normalizationTargetPidChecks++
                    $script:normalizationCompletedStatus |
                        ConvertTo-Json |
                        Set-Content -LiteralPath $script:normalizationStatusPath -Encoding UTF8
                    return [PSCustomObject]@{ Id = 12345 }
                }
                if ($script:normalizationIpcStarted) {
                    return @(
                        [PSCustomObject]@{ Id = 12345 },
                        [PSCustomObject]@{ Id = 54321 }
                    )
                }
                return [PSCustomObject]@{ Id = 12345 }
            }
            Mock Start-Process {
                param(
                    $FilePath,
                    $ArgumentList,
                    $WindowStyle,
                    $RedirectStandardOutput,
                    $RedirectStandardError,
                    $PassThru
                )
                $null = @(
                    $FilePath,
                    $WindowStyle,
                    $RedirectStandardOutput,
                    $RedirectStandardError,
                    $PassThru
                )
                $script:normalizationIpcArguments = @($ArgumentList)
                $ipcScriptPath = $ArgumentList[0].Substring('--execute='.Length).Trim('"')
                $ipcScriptText = Get-Content -LiteralPath $ipcScriptPath -Raw -Encoding UTF8
                $manifestMatch = [regex]::Match(
                    $ipcScriptText,
                    'var ASITOPIX_MANIFEST_PATH = (?<literal>".*?");'
                )
                $manifestPath = $manifestMatch.Groups["literal"].Value | ConvertFrom-Json
                $manifest = Get-Content -LiteralPath $manifestPath -Raw -Encoding UTF8 | ConvertFrom-Json
                foreach ($frame in $manifest.frames) {
                    Set-Content -LiteralPath $frame.normalizedPath -Value "normalized"
                    Set-Content -LiteralPath $frame.colorCorrectedPath -Value "normalized and color corrected"
                }
                Set-Content -LiteralPath $manifest.metricsPath -Value "Block,k"
                $script:normalizationCompletedStatus = [ordered]@{
                    schemaVersion         = 1
                    state                 = "completed"
                    currentBlockNumber    = $null
                    completedBlockNumbers = @($manifest.frames.blockNumber)
                    errorMessage          = ""
                    errorStack            = ""
                }
                [ordered]@{
                    schemaVersion         = 1
                    state                 = "running"
                    currentBlockNumber    = 1
                    completedBlockNumbers = @()
                    errorMessage          = ""
                    errorStack            = ""
                } | ConvertTo-Json | Set-Content -LiteralPath $manifest.statusPath -Encoding UTF8
                $script:normalizationStatusPath = $manifest.statusPath
                $script:normalizationIpcStarted = $true

                $process = [PSCustomObject]@{
                    ExitCode = 0
                    HasExited = $true
                }
                $process | Add-Member -MemberType ScriptMethod -Name Refresh -Value { }
                $process | Add-Member -MemberType ScriptMethod -Name WaitForExit -Value { }
                return $process
            }

            $result = Invoke-AsiToPixTotalityNormalizationPlan `
                -Plan $plan `
                -MetricsPath $metricsPath `
                -PixInsightPath $fakePixInsightPath `
                -PixInsightScriptPath $scriptPath `
                -Confirm:$false

            $result.NormalizedCount | Should Be 1
            Test-Path -LiteralPath $plan[0].NormalizedPath -PathType Leaf | Should Be $true
            Test-Path -LiteralPath $plan[0].ColorCorrectedPath -PathType Leaf | Should Be $true
            Test-Path -LiteralPath $metricsPath -PathType Leaf | Should Be $true
            Get-Content -LiteralPath $plan[0].NormalizedPath | Should Be "normalized"
            Get-Content -LiteralPath $plan[0].ColorCorrectedPath | Should Be "normalized and color corrected"
            Get-Content -LiteralPath $metricsPath | Should Be "Block,k"
            $script:normalizationIpcArguments.Count | Should Be 1
            $script:normalizationIpcArguments[0] | Should Match '^--execute='
            $script:normalizationTargetPidChecks | Should Be 1
            Assert-MockCalled Start-Process -Times 1
        }

        Remove-Item Env:\ASITOPIX_TOTALITY_NORMALIZATION_TEST_ROOT
    }
}

Describe "PixInsight Totality normalization implementation" {
    It "uses the active preview ROI and measures each RGB channel median" {
        $scriptText = Get-Content -LiteralPath $pixInsightScriptPath -Raw -Encoding UTF8

        $scriptText | Should Match 'window\.currentView'
        $scriptText | Should Match 'window\.previews'
        $scriptText | Should Match 'window\.previewRect\( preview \)'
        $scriptText | Should Match 'image\.selectedRect = rect;'
        $scriptText | Should Match 'image\.selectedChannel = channel;'
        $scriptText | Should Match 'var median = image\.median\(\);'
    }

    It "implements the requested scalar and fixed color formulas" {
        $scriptText = Get-Content -LiteralPath $pixInsightScriptPath -Raw -Encoding UTF8

        $scriptText | Should Match 'referenceMedians\[0\] / medians\[0\]'
        $scriptText | Should Match 'referenceMedians\[1\] / medians\[1\]'
        $scriptText | Should Match 'referenceMedians\[2\] / medians\[2\]'
        $scriptText | Should Match 'var scalar = median3\( factors\[0\], factors\[1\], factors\[2\] \);'
        $scriptText | Should Match 'view\.image\.apply\( scalar, ImageOp_Mul \);'
        $scriptText | Should Match 'referenceMedians\[1\] / referenceMedians\[0\]'
        $scriptText | Should Match 'referenceMedians\[1\] / referenceMedians\[2\]'
        $scriptText | Should Match 'view\.image\.apply\( coefficients\[channel\], ImageOp_Mul \);'
    }

    It "saves both image variants and the CSV metrics through the IPC manifest" {
        $scriptText = Get-Content -LiteralPath $pixInsightScriptPath -Raw -Encoding UTF8

        $scriptText | Should Match 'window\.saveAs\( frame\.normalizedPath, false, false, false, false \);'
        $scriptText | Should Match 'window\.saveAs\( frame\.colorCorrectedPath, false, false, false, false \);'
        $scriptText | Should Match 'File\.writeTextFile\( manifest\.metricsPath, csvLines\.join'
        $scriptText | Should Not Match 'Normalization metrics CSV already exists'
        $scriptText | Should Not Match 'Normalization output already exists'
        $scriptText | Should Not Match 'Color-corrected output already exists'
        $scriptText | Should Match '"Rref", "Gref", "Bref", "R", "G", "B"'
        $scriptText | Should Match 'File\.writeTextFile\( statusPath, JSON\.stringify\( status, null, 2 \) \);'
        $scriptText | Should Match '// ASITOPIX_IPC_MANIFEST'
    }

    It "prompts for an opened reference and preview before IPC execution" {
        $entryPointText = Get-Content -LiteralPath $entryPointPath -Raw -Encoding UTF8

        $entryPointText | Should Match 'ChildPath "Composed"'
        $entryPointText | Should Match 'ChildPath "Calibrated"'
        $entryPointText | Should Match 'Open one Block-\*-HDR\.xisf frame'
        $entryPointText | Should Match 'Create a preview over the region'
        $entryPointText | Should Match 'Read-Host "Press Enter here when the reference frame and preview are ready"'
    }
}
