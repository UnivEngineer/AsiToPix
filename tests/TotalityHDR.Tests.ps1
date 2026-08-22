$modulePath = Join-Path -Path $PSScriptRoot -ChildPath "..\src\AsiToPix.TotalityHDR.psm1"
Import-Module $modulePath -Force

function Add-TestTotalityHdrFrame {
    param(
        [string]$Directory,
        [int]$Block,
        [string]$Exposure,
        [string]$Suffix = "Totality_00123_00150_ca_r180"
    )

    $path = Join-Path -Path $Directory -ChildPath (
        "Block_{0:D3}_{1}_{2}.xisf" -f $Block, $Exposure, $Suffix
    )
    Set-Content -LiteralPath $path -Value "test"
    return $path
}

Describe "Totality HDR frame discovery" {
    It "extracts block numbers and exposures from aligned XISF names" {
        $inputPath = Join-Path -Path $TestDrive -ChildPath "hdr-discovery"
        New-Item -ItemType Directory -Path $inputPath -Force | Out-Null
        Add-TestTotalityHdrFrame -Directory $inputPath -Block 9 -Exposure "10ms" | Out-Null
        Add-TestTotalityHdrFrame -Directory $inputPath -Block 5 -Exposure "100ms" | Out-Null
        Set-Content -LiteralPath (Join-Path $inputPath "notes.txt") -Value "ignore"

        $frames = @(Get-AsiToPixTotalityHdrFrame -InputPath $inputPath)

        $frames.Count | Should Be 2
        $frames[0].BlockNumber | Should Be 5
        $frames[0].ExposureLabel | Should Be "100ms"
        $frames[1].BlockNumber | Should Be 9
        $frames[1].ExposureMilliseconds | Should Be ([decimal]10)
    }

    It "rejects two files for the same numeric block and exposure" {
        $inputPath = Join-Path -Path $TestDrive -ChildPath "hdr-duplicates"
        New-Item -ItemType Directory -Path $inputPath -Force | Out-Null
        Set-Content -LiteralPath (Join-Path $inputPath "Block_005_100ms_first.xisf") -Value "first"
        Set-Content -LiteralPath (Join-Path $inputPath "Block_005_100ms_second.xisf") -Value "second"

        @(Get-ChildItem -LiteralPath $inputPath -File).Count | Should Be 2
        $caughtError = $null
        try {
            $null = @(Get-AsiToPixTotalityHdrFrame -InputPath $inputPath)
        }
        catch {
            $caughtError = $_
        }
        $caughtError | Should Not BeNullOrEmpty
    }
}

Describe "Totality HDR exposure ladders" {
    It "offers every contiguous ladder for three discovered exposures" {
        $inputPath = Join-Path -Path $TestDrive -ChildPath "hdr-ladders"
        New-Item -ItemType Directory -Path $inputPath -Force | Out-Null
        foreach ($block in @(1, 2)) {
            foreach ($exposure in @("1ms", "10ms", "100ms")) {
                Add-TestTotalityHdrFrame -Directory $inputPath -Block $block -Exposure $exposure | Out-Null
            }
        }
        $frames = @(Get-AsiToPixTotalityHdrFrame -InputPath $inputPath)

        $ladders = @(Get-AsiToPixTotalityHdrLadder -Frames $frames)

        $ladders.Count | Should Be 3
        $ladders[0].Label | Should Be "1ms + 10ms + 100ms"
        $ladders[1].Label | Should Be "1ms + 10ms"
        $ladders[2].Label | Should Be "10ms + 100ms"
        $ladders[0].CompleteBlockCount | Should Be 2
    }

    It "selects a requested available ladder independent of argument order" {
        $inputPath = Join-Path -Path $TestDrive -ChildPath "hdr-selection"
        New-Item -ItemType Directory -Path $inputPath -Force | Out-Null
        foreach ($exposure in @("1ms", "10ms", "100ms")) {
            Add-TestTotalityHdrFrame -Directory $inputPath -Block 1 -Exposure $exposure | Out-Null
        }
        $frames = @(Get-AsiToPixTotalityHdrFrame -InputPath $inputPath)

        $selected = Select-AsiToPixTotalityHdrLadder `
            -Frames $frames `
            -Exposure @("100ms", "10ms")

        $selected.Label | Should Be "10ms + 100ms"
    }
}

Describe "Totality HDR plan" {
    It "orders each complete block from longest to shortest exposure" {
        $inputPath = Join-Path -Path $TestDrive -ChildPath "hdr-plan-input"
        $outputPath = Join-Path -Path $TestDrive -ChildPath "hdr-plan-output"
        New-Item -ItemType Directory -Path $inputPath -Force | Out-Null
        Add-TestTotalityHdrFrame -Directory $inputPath -Block 5 -Exposure "10ms" | Out-Null
        Add-TestTotalityHdrFrame -Directory $inputPath -Block 5 -Exposure "100ms" | Out-Null
        Add-TestTotalityHdrFrame -Directory $inputPath -Block 6 -Exposure "100ms" | Out-Null
        $frames = @(Get-AsiToPixTotalityHdrFrame -InputPath $inputPath)
        $ladder = Select-AsiToPixTotalityHdrLadder -Frames $frames

        $plan = @(Get-AsiToPixTotalityHdrPlan `
            -Frames $frames `
            -Ladder $ladder `
            -OutputDirectory $outputPath)

        $plan.Count | Should Be 2
        $plan[0].CanCompose | Should Be $true
        $plan[0].Frames[0].ExposureLabel | Should Be "100ms"
        $plan[0].Frames[1].ExposureLabel | Should Be "10ms"
        $plan[0].OutputPath | Should Be (Join-Path $outputPath "Block-005-HDR.xisf")
        $plan[1].CanCompose | Should Be $false
        $plan[1].SkipReason | Should Match "10ms"
    }
}

Describe "PixInsight Totality HDR script" {
    It "loads the HDRComposition icon, assigns manifest images, saves, and closes" {
        $scriptPath = Join-Path -Path $PSScriptRoot -ChildPath "..\pixinsight\TotalityHDR.js"
        $scriptText = Get-Content -LiteralPath $scriptPath -Raw -Encoding UTF8

        $scriptText | Should Match 'ProcessInstance\.fromIcon\( "HDRComposition" \)'
        $scriptText | Should Match 'if \( P === null \)'
        $scriptText | Should Match 'P = new HDRComposition;'
        $scriptText | Should Match 'images\.push\( \[ true, block\.inputFiles\[i\] \] \)'
        $scriptText | Should Match 'P\.images = images;'
        $scriptText | Should Match 'P\.executeGlobal\(\)'
        $scriptText | Should Match 'ImageWindow\.activeWindow'
        $scriptText | Should Match 'outputWindow\.saveAs\( block\.outputPath, false, false, false, false \);'
        $scriptText | Should Match 'outputWindow\.forceClose\(\);'
        $scriptText | Should Match '// ASITOPIX_IPC_MANIFEST'
    }
}

Describe "Totality HDR entry point" {
    It "builds a WhatIf plan without creating the HDR output directory" {
        $astroPhotoRoot = Join-Path -Path $TestDrive -ChildPath "AstroPhoto"
        $inputPath = Join-Path -Path $astroPhotoRoot -ChildPath "SharpCap\Totality-2026\HDR\Aligned"
        $outputPath = Join-Path -Path $astroPhotoRoot -ChildPath "SharpCap\Totality-2026\HDR\Composed"
        $pixInsightPath = Join-Path -Path $TestDrive -ChildPath "PixInsight.exe"
        New-Item -ItemType Directory -Path $inputPath -Force | Out-Null
        Add-TestTotalityHdrFrame -Directory $inputPath -Block 1 -Exposure "10ms" | Out-Null
        Add-TestTotalityHdrFrame -Directory $inputPath -Block 1 -Exposure "100ms" | Out-Null
        Set-Content -LiteralPath $pixInsightPath -Value "test executable"
        $entryScript = Join-Path -Path $PSScriptRoot -ChildPath "..\HDRTotality.ps1"

        $null = & $entryScript `
            -AstroPhotoRoot $astroPhotoRoot `
            -Exposure @("10ms", "100ms") `
            -PixInsightPath $pixInsightPath `
            -PixInsightMode Dedicated `
            -WhatIf `
            -Confirm:$false

        Test-Path -LiteralPath $outputPath | Should Be $false
    }

    It "uses shared root aliases and both PixInsight execution modes" {
        $entryScript = Join-Path -Path $PSScriptRoot -ChildPath "..\HDRTotality.ps1"
        $scriptText = Get-Content -LiteralPath $entryScript -Raw -Encoding UTF8

        $scriptText | Should Match 'Resolve-AstroPhotoRoot'
        $scriptText | Should Match 'SharpCap\\Totality-2026\\HDR'
        $scriptText | Should Match 'ChildPath "Composed"'
        $scriptText | Should Match '\[ValidateSet\("Reuse", "Dedicated"\)\]'
        $scriptText | Should Match 'Read-AsiToPixPixInsightMode'
        $scriptText | Should Match 'Operation\s+= "HDR"'
    }
}
